import logging
import threading
from typing import Callable, Mapping, Optional, Tuple

try:
    import tkinter as tk
except Exception:  # pragma: no cover - tkinter may be missing on headless hosts
    tk = None  # type: ignore

LogCallback = Callable[[str, str, Optional[str]], None]


class OverlayController:
    """Encapsulates the Tk overlay so it can be reused and tested."""

    def __init__(
        self,
        pause_event: threading.Event,
        stop_event: threading.Event,
        logger: logging.Logger,
        log_callback: Optional[LogCallback] = None,
        color_map: Optional[Mapping[str, str]] = None,
    ) -> None:
        self.pause_event = pause_event
        self.stop_event = stop_event
        self.logger = logger
        self.log_callback = log_callback
        self.color_map = dict(color_map) if color_map else {}
        self._thread: Optional[threading.Thread] = None

    @property
    def available(self) -> bool:
        return tk is not None

    def launch(self, window_bounds: Optional[Tuple[int, int, int, int]] = None) -> Optional[threading.Thread]:
        """Start the overlay in a daemon thread if tkinter is present."""
        if not self.available:
            self.logger.warning('tkinter is not available. Overlay controls disabled.')
            return None
        if self._thread and self._thread.is_alive():
            return self._thread
        self._thread = threading.Thread(
            target=self._run_overlay,
            args=(window_bounds,),
            name='OverlayUI',
            daemon=True,
        )
        self._thread.start()
        return self._thread

    def _log(self, level: str, message: str, color: Optional[str] = None) -> None:
        if self.log_callback:
            self.log_callback(level, message, color)
        else:
            log_fn = getattr(self.logger, level, self.logger.info)
            log_fn(message)

    def _run_overlay(self, window_bounds: Optional[Tuple[int, int, int, int]]) -> None:
        if tk is None:
            return
        try:
            root = tk.Tk()
        except Exception:
            self.logger.exception('Failed to initialize overlay window')
            return

        root.title('FH5 Sniper Control')
        root.configure(bg='#111111')
        root.attributes('-topmost', True)
        root.attributes('-alpha', 0.9)
        root.overrideredirect(True)

        status_var = tk.StringVar(value='Running')

        def place_overlay(bounds: Optional[Tuple[int, int, int, int]]):
            if not bounds:
                root.geometry('+80+80')
                return
            left, top, width, height = bounds
            margin = 16
            root.update_idletasks()
            ow = root.winfo_width() or 180
            oh = root.winfo_height() or 120
            x = int(left + width - ow - margin)
            y = int(top + height - oh - margin)
            min_x = int(left + margin)
            min_y = int(top + margin)
            root.geometry(f'+{max(x, min_x)}+{max(y, min_y)}')

        place_overlay(window_bounds)

        def toggle_pause():
            if self.pause_event.is_set():
                self.pause_event.clear()
                status_var.set('Running')
                pause_btn.configure(text='Pause')
                self._log('info', 'Automation resumed from overlay', self.color_map.get('resume'))
            else:
                self.pause_event.set()
                status_var.set('Paused')
                pause_btn.configure(text='Resume')
                self._log('warning', 'Automation paused from overlay', self.color_map.get('pause'))

        def request_stop():
            if not self.stop_event.is_set():
                self.stop_event.set()
                status_var.set('Stopping...')
                self._log('warning', 'Stop requested from overlay', self.color_map.get('stop'))
            try:
                root.destroy()
            except tk.TclError:
                pass

        def start_move(event):
            root._drag_start_x = event.x  # type: ignore[attr-defined]
            root._drag_start_y = event.y  # type: ignore[attr-defined]

        def do_move(event):
            x = root.winfo_pointerx() - getattr(root, '_drag_start_x', 0)
            y = root.winfo_pointery() - getattr(root, '_drag_start_y', 0)
            root.geometry(f'+{x}+{y}')

        frame = tk.Frame(root, bg='#111111', padx=12, pady=10)
        frame.pack()

        title_lbl = tk.Label(frame, text='Sniper Overlay', fg='white', bg='#111111', font=('Segoe UI', 10, 'bold'))
        title_lbl.pack(anchor='w')
        status_lbl = tk.Label(frame, textvariable=status_var, fg='#66ff99', bg='#111111', font=('Consolas', 10))
        status_lbl.pack(anchor='w', pady=(2, 8))

        btn_style = {
            'bg': '#1e1e1e',
            'fg': 'white',
            'activebackground': '#4b4b4b',
            'activeforeground': 'white',
            'bd': 0,
            'font': ('Segoe UI', 10, 'bold'),
            'width': 12,
            'pady': 4,
        }

        pause_btn = tk.Button(frame, text='Pause', command=toggle_pause, **btn_style)
        pause_btn.pack(pady=(0, 6))
        stop_btn = tk.Button(
            frame,
            text='Stop',
            command=request_stop,
            bg='#8b0000',
            activebackground='#a40000',
            fg='white',
            width=12,
            pady=4,
            bd=0,
            font=('Segoe UI', 10, 'bold'),
        )
        stop_btn.pack()

        for widget in (frame, title_lbl, status_lbl):
            widget.bind('<Button-1>', start_move)
            widget.bind('<B1-Motion>', do_move)

        def monitor_stop_flag():
            if self.stop_event.is_set():
                try:
                    root.destroy()
                except tk.TclError:
                    pass
            else:
                root.after(200, monitor_stop_flag)

        root.bind('<Escape>', lambda _event: request_stop())
        root.after(200, monitor_stop_flag)
        root.mainloop()
