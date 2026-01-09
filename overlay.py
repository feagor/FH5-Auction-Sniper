import logging
import threading
import time
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
        refocus_callback: Optional[Callable[[], None]] = None,
    ) -> None:
        self.pause_event = pause_event
        self.stop_event = stop_event
        self.logger = logger
        self.log_callback = log_callback
        self.color_map = dict(color_map) if color_map else {}
        self._refocus_callback = refocus_callback
        self._thread: Optional[threading.Thread] = None
        self._info_lock = threading.Lock()
        self._current_car = '—'
        self._remaining_buyouts = 0
        self._purchased_count = 0
        self._session_purchases = 0
        self._remaining_seconds: float = 0.0
        self._last_tick: float = time.time()

    @property
    def available(self) -> bool:
        return tk is not None

    def update_status(
        self,
        car_name: Optional[str] = None,
        remaining_seconds: Optional[float] = None,
        remaining_buyouts: Optional[int] = None,
        purchased_count: Optional[int] = None,
        session_purchases: Optional[int] = None,
    ) -> None:
        """Store overlay info from any thread; picked up by UI poller."""
        with self._info_lock:
            if car_name is not None:
                self._current_car = car_name or '—'
            if remaining_seconds is not None:
                self._remaining_seconds = max(0.0, float(remaining_seconds))
                self._last_tick = time.time()
            if remaining_buyouts is not None:
                self._remaining_buyouts = max(0, int(remaining_buyouts))
            if purchased_count is not None:
                self._purchased_count = max(0, int(purchased_count))
            if session_purchases is not None:
                self._session_purchases = max(0, int(session_purchases))

    def get_remaining_seconds(self) -> int:
        """Expose the current countdown so the main loop can stay in sync."""
        with self._info_lock:
            remaining = self._update_remaining_locked(time.time())
            return int(remaining)

    def _update_remaining_locked(self, now: float) -> float:
        remaining = self._remaining_seconds
        if remaining <= 0:
            self._remaining_seconds = 0.0
            self._last_tick = now
            return 0.0
        if self.pause_event.is_set():
            return remaining
        elapsed = max(0.0, now - self._last_tick)
        if elapsed > 0:
            remaining = max(0.0, remaining - elapsed)
            self._remaining_seconds = remaining
            self._last_tick = now
        return remaining

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
        root.minsize(230, 210)

        status_var = tk.StringVar(value='Running')

        def place_overlay(bounds: Optional[Tuple[int, int, int, int]]):
            if not bounds:
                root.geometry('+80+80')
                return
            left, top, width, height = bounds
            margin = 16
            root.update_idletasks()
            ow = max(root.winfo_width() or 0, 230)
            oh = root.winfo_height() or 120
            x = int(left + width - ow - margin)
            y = int(top + height - oh - margin)
            min_x = int(left + margin)
            min_y = int(top + margin)
            root.geometry(f'+{max(x - 50, min_x)}+{max(y - 150, min_y)}')

        def toggle_pause():
            if self.pause_event.is_set():
                self.pause_event.clear()
                with self._info_lock:
                    self._last_tick = time.time()
                status_var.set('Running')
                pause_btn.configure(text='Pause')
                self._log('info', 'Automation resumed from overlay', self.color_map.get('resume'))
                self._refocus_if_needed()
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

        frame = tk.Frame(root, bg='#111111', padx=12, pady=10)
        frame.pack()

        title_lbl = tk.Label(frame, text='Sniper Overlay', fg='white', bg='#111111', font=('Segoe UI', 10, 'bold'))
        title_lbl.pack(anchor='w')
        status_lbl = tk.Label(frame, textvariable=status_var, fg='#66ff99', bg='#111111', font=('Consolas', 10))
        status_lbl.pack(anchor='w', pady=(2, 8))

        car_var = tk.StringVar(value='Car: —')
        timer_var = tk.StringVar(value='Window: 00:00')
        stock_var = tk.StringVar(value='Left: 0')
        bought_var = tk.StringVar(value='Hits: 0')
        session_var = tk.StringVar(value='Session Hits: 0')
        car_lbl = tk.Label(frame, textvariable=car_var, fg='white', bg='#111111', font=('Consolas', 10))
        timer_lbl = tk.Label(frame, textvariable=timer_var, fg='#66ccff', bg='#111111', font=('Consolas', 10))
        stock_lbl = tk.Label(frame, textvariable=stock_var, fg='#ffcc66', bg='#111111', font=('Consolas', 10))
        bought_lbl = tk.Label(frame, textvariable=bought_var, fg='#66ffcc', bg='#111111', font=('Consolas', 10))
        session_lbl = tk.Label(frame, textvariable=session_var, fg='#f2a0ff', bg='#111111', font=('Consolas', 10))
        car_lbl.pack(anchor='w')
        timer_lbl.pack(anchor='w')
        stock_lbl.pack(anchor='w')
        bought_lbl.pack(anchor='w')
        session_lbl.pack(anchor='w', pady=(0, 8))

        place_overlay(window_bounds)

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
       
        def refresh_overlay_info():
            now = time.time()
            with self._info_lock:
                car_name = self._current_car
                remaining_buyouts = self._remaining_buyouts
                purchased = self._purchased_count
                session_total = self._session_purchases
                remaining = self._update_remaining_locked(now)
            remaining_secs = int(round(remaining))
            minutes = remaining_secs // 60
            seconds = remaining_secs % 60
            car_var.set(f'Target: {car_name or "—"}')
            timer_var.set(f'Window: {minutes:02d}:{seconds:02d}')
            stock_var.set(f'Left: {remaining_buyouts}')
            bought_var.set(f'Car Hits: {purchased}')
            session_var.set(f'Session Hits: {session_total}')
            root.after(1000, refresh_overlay_info)

        def monitor_stop_flag():
            if self.stop_event.is_set():
                try:
                    root.destroy()
                except tk.TclError:
                    pass
            else:
                root.after(200, monitor_stop_flag)

        pause_state = self.pause_event.is_set()
        if pause_state:
            status_var.set('Paused')
            pause_btn.configure(text='Resume')

        def monitor_pause_flag():
            nonlocal pause_state
            is_paused = self.pause_event.is_set()
            if is_paused != pause_state:
                pause_state = is_paused
                if is_paused:
                    status_var.set('Paused')
                    pause_btn.configure(text='Resume')
                else:
                    status_var.set('Running')
                    pause_btn.configure(text='Pause')
                    with self._info_lock:
                        self._last_tick = time.time()
                    self._refocus_if_needed()
            root.after(200, monitor_pause_flag)

        root.after(200, refresh_overlay_info)
        root.after(200, monitor_stop_flag)
        root.after(200, monitor_pause_flag)
        root.mainloop()

    def _refocus_if_needed(self) -> None:
        if not self._refocus_callback:
            return
        try:
            self._refocus_callback()
        except Exception:
            self.logger.exception('Failed to refocus target window after resume')
