from ctypes import windll, byref, c_int
import win32api
import win32con
import wmi

PROCESS_PER_MONITOR_DPI_AWARE = 2
MDT_EFFECTIVE_DPI = 0

def get_monitors():	
	windll.shcore.SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE)
	# Collect friendly names from WMI (root\wmi\WmiMonitorID)
	wmi_names = {}
	try:
		w = wmi.WMI(namespace="root\\wmi")
		for m in w.WmiMonitorID():
			if m.UserFriendlyName is None:
				continue
			try:
				parts = m.InstanceName.split("\\")
				if len(parts) > 1:
					pnp_id = parts[1]
					name = "".join(chr(x) for x in m.UserFriendlyName if x != 0)
					wmi_names[pnp_id] = name
			except Exception:
				continue
	except Exception:
		pass

	# Helper: get scale (%) for a monitor handle
	def _get_scale_factor(h_monitor):
		try:
			dpi_x = c_int()
			dpi_y = c_int()
			res = windll.shcore.GetDpiForMonitor(
				h_monitor,
				MDT_EFFECTIVE_DPI,
				byref(dpi_x),
				byref(dpi_y)
			)
			if res == 0:
				return int((dpi_x.value / 96) * 100)
		except Exception:
			pass
		return 100

	result = []
	monitors = win32api.EnumDisplayMonitors()

	for h_monitor_py, _, _ in monitors:
		h_monitor = int(h_monitor_py)

		monitor_data = {
			"name": "Generic Monitor",
			"resolution": "Unknown",
			"refresh_rate": 0,
			"scale": 100,
			"handle": h_monitor,
			# geometry (left, top, right, bottom)
			"left": 0,
			"top": 0,
			"right": 0,
			"bottom": 0,
		}

		# 1. Get adapter/device name (e.g. \\.\DISPLAY1)
		try:
			info = win32api.GetMonitorInfo(h_monitor)
			device_name = info['Device']
			try:
				mon_rect = info.get('Monitor')  # (left, top, right, bottom)
				if mon_rect and len(mon_rect) >= 4:
					l, t, r, b = map(int, mon_rect[:4])
					monitor_data['left'] = l
					monitor_data['top'] = t
					monitor_data['right'] = r
					monitor_data['bottom'] = b
			except Exception:
				pass
		except Exception:
			continue

		# 2. Get resolution and refresh rate
		try:
			settings = win32api.EnumDisplaySettings(device_name, win32con.ENUM_CURRENT_SETTINGS)
			monitor_data["resolution"] = f"{settings.PelsWidth}x{settings.PelsHeight}"
			monitor_data["refresh_rate"] = settings.DisplayFrequency
		except Exception:
			pass

		# 3. Get scale (DPI)
		monitor_data["scale"] = _get_scale_factor(h_monitor)

		# 4. Get friendly name from WMI if available, otherwise system DeviceString
		try:
			mon_device = win32api.EnumDisplayDevices(device_name, 0)
			if "\\" in mon_device.DeviceID:
				pnp_id = mon_device.DeviceID.split("\\")[1]
				monitor_data["name"] = wmi_names.get(pnp_id, mon_device.DeviceString)
		except Exception:
			pass
		result.append(monitor_data)
	return result

def find_monitor_for_point(x, y, monitors_list):
	"""
	Return monitor dict that contains point (x,y) or None.
	"""
	for m in monitors_list:
		l = int(m.get('left', 0))
		t = int(m.get('top', 0))
		r = int(m.get('right', l))
		b = int(m.get('bottom', t))
		if l <= x < r and t <= y < b:
			return m
	return None

if __name__ == "__main__":
	monitors_list = get_monitors()
	print(f"Monitors found: {len(monitors_list)}\n")
	for m in monitors_list:
		print(f"Monitor:    {m['name']}")
		print(f"Resolution: {m['resolution']}")
		print(f"Refresh:    {m['refresh_rate']} Hz")
		print(f"Scale:      {m['scale']}%")
		print(f"Handle:     {m['handle']} (HEX: {hex(m['handle'])})")
		print(f"Geometry:   Left={m['left']}, Top={m['top']}, Right={m['right']}, Bottom={m['bottom']}")
		print("-" * 30)