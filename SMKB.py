import threading
import time
import random
import sys
import re
import math
import tkinter as tk
from tkinter import messagebox
from tkinter import StringVar, IntVar
from tkinter import ttk

try:
    from pynput import keyboard, mouse
    from pynput.keyboard import Key, Controller as KController
    from pynput.mouse import Button as MouseButton, Controller as MController
except Exception as e:
    print("Missing dependency 'pynput'. Install with: pip install pynput")
    raise

def ms_to_sec(ms):
    return max(0.0, ms) / 1000.0

def clamp(v, a, b):
    return max(a, min(b, v))

def parse_key_name(name: str):
    name = name.strip().lower()
    if not name:
        return None
    mapping = {
        'enter': Key.enter, 'return': Key.enter, 'space': Key.space, 'tab': Key.tab,
        'esc': Key.esc, 'escape': Key.esc, 'backspace': Key.backspace,
        'shift': Key.shift, 'ctrl': Key.ctrl, 'control': Key.ctrl, 'alt': Key.alt,
        'cmd': Key.cmd, 'super': Key.cmd, 'win': Key.cmd, 'capslock': Key.caps_lock,
        'up': Key.up, 'down': Key.down, 'left': Key.left, 'right': Key.right,
        'home': Key.home, 'end': Key.end, 'pageup': Key.page_up, 'pagedown': Key.page_down,
        'insert': Key.insert, 'delete': Key.delete,
        'f1': Key.f1, 'f2': Key.f2, 'f3': Key.f3, 'f4': Key.f4,
        'f5': Key.f5, 'f6': Key.f6, 'f7': Key.f7, 'f8': Key.f8,
        'f9': Key.f9, 'f10': Key.f10, 'f11': Key.f11, 'f12': Key.f12,
    }
    if name in mapping:
        return mapping[name]
    if len(name) == 1:
        return name
    return name

class AutoController:
    def __init__(self):
        self.kc = KController()
        self.mc = MController()
        self.running = False
        self._thread = None
        self._stop_event = threading.Event()

    def start(self, job_fn):
        if self.running:
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run_loop, args=(job_fn,), daemon=True)
        self.running = True
        self._thread.start()

    def _run_loop(self, job_fn):
        try:
            job_fn(stop_event=self._stop_event)
        except Exception as e:
            print('Error in automation thread:', e)
        finally:
            self.running = False

    def stop(self):
        if not self.running:
            return
        self._stop_event.set()
        if self._thread:
            self._thread.join(timeout=1)
        self.running = False

MODIFIER_RE = re.compile(r"(\w+)=([-\w.,]+)")

def parse_sequence(raw: str):
    actions = []
    if not raw:
        return actions
    items = [s.strip() for s in raw.split(',') if s.strip()]
    for it in items:
        parts = [p.strip() for p in it.split('|') if p.strip()]
        if not parts: continue
        key_part = parts[0]
        modifiers = parts[1:]
        key_names = [k.strip() for k in key_part.split('+') if k.strip()]
        parsed_keys = [parse_key_name(k) for k in key_names]
        simul = len(parsed_keys) > 1
        hold = None; repeat = 1
        for m in modifiers:
            mm = MODIFIER_RE.match(m)
            if not mm: continue
            k, v = mm.group(1).lower(), mm.group(2)
            if k == 'hold':
                try: hold = int(v)
                except: hold = None
            elif k == 'repeat':
                try: repeat = max(1, int(v))
                except: repeat = 1
        actions.append({'keys': parsed_keys, 'simul': simul, 'hold': hold, 'repeat': repeat, 'raw': it})
    return actions

def parse_chant(raw: str):
    steps = []
    if not raw:
        return steps
    raw_steps = [s.strip() for s in raw.split(';') if s.strip()]
    for st in raw_steps:
        parts = [p.strip() for p in st.split('||') if p.strip()]
        step_actions = []
        for p in parts:
            low = p.lower().strip()
            if low == 'stop':
                step_actions.append({'device':'stop', 'raw':p})
                continue
            if p.lower().startswith('mouse') or p.lower().startswith('m(') or p.lower().startswith('m '):
                tok_parts = [q.strip() for q in p.split('|') if q.strip()]
                pos = None
                move = None; hold = None; button = 'left'; rel = None; dist = None; repeat = 1
                first = tok_parts[0]
                mcoords = re.match(r'm(?:ouse)?\s*\(\s*([-\d]+)\s*,\s*([-\d]+)\s*\)', first, re.I)
                if mcoords:
                    try:
                        pos = (int(mcoords.group(1)), int(mcoords.group(2)))
                    except:
                        pos = None
                for mod in tok_parts[1:]:
                    mm = MODIFIER_RE.match(mod)
                    if not mm: continue
                    k, v = mm.group(1).lower(), mm.group(2)
                    if k == 'move':
                        try: move = int(v)
                        except: pass
                    elif k == 'hold':
                        try: hold = int(v)
                        except: pass
                    elif k == 'button':
                        if v.lower() in ('left','right'):
                            button = v.lower()
                    elif k == 'rel':
                        try: rel = float(v)
                        except: pass
                    elif k == 'dist':
                        try: dist = int(v)
                        except: pass
                    elif k == 'repeat':
                        try: repeat = max(1, int(v))
                        except: pass
                step_actions.append({'device':'mouse', 'pos':pos, 'move':move, 'hold':hold, 'button':button, 'rel':rel, 'dist':dist, 'repeat':repeat, 'raw':p})
            else:
                kb_parsed = parse_sequence(p)
                if not kb_parsed:
                    continue
                for kact in kb_parsed:
                    kact_out = dict(kact)
                    kact_out['device'] = 'kb'
                    step_actions.append(kact_out)
        if step_actions:
            steps.append(step_actions)
    return steps

def ease_out_cubic(t: float) -> float:
    return 1 - pow(1 - t, 3)

def ease_in_out_cubic(t: float) -> float:
    if t < 0.5: return 4 * t * t * t
    else: return 1 - pow(-2 * t + 2, 3) / 2

class App:
    def __init__(self, master):
        self.master = master
        master.title('Auto Input Controller — Chant + Ribbon')
        self.controller = AutoController()
        self._pressed_keys = set()
        self._pressed_buttons = set()
        self.style = ttk.Style(master)
        try:
            self.style.theme_use('clam')
        except Exception:
            pass
        master.configure(bg='#2b2b2b')
        self.style.configure('.', background='#2b2b2b', foreground='#e6e6e6', font=('Segoe UI', 10), relief='flat')
        self.style.configure('TFrame', background='#2b2b2b')
        self.style.configure('TLabelFrame', background='#2b2b2b', foreground='#e6e6e6')
        self.style.configure('TLabel', background='#2b2b2b', foreground='#e6e6e6')
        self.style.configure('TEntry', fieldbackground='#3a3a3a', background='#3a3a3a', foreground='#ffffff')
        self.style.configure('TButton', background='#444444', foreground='#ffffff')
        self.style.map('TButton', background=[('active', '#555555'), ('pressed', '#333333')])
        pad = {'padx': 6, 'pady': 6}
        notebook = ttk.Notebook(master)
        notebook.grid(row=0, column=0, sticky='nsew', padx=6, pady=6)
        tab_home = ttk.Frame(notebook)
        tab_kb = ttk.Frame(notebook)
        tab_mouse = ttk.Frame(notebook)
        notebook.add(tab_home, text='Home')
        notebook.add(tab_kb, text='Keyboard')
        notebook.add(tab_mouse, text='Mouse')
        master.grid_rowconfigure(0, weight=1)
        master.grid_columnconfigure(0, weight=1)
        tab_home.grid_columnconfigure(0, weight=1)
        tab_kb.grid_columnconfigure(0, weight=1)
        tab_mouse.grid_columnconfigure(0, weight=1)
        hk_frame = ttk.LabelFrame(tab_home, text='Hotkey')
        hk_frame.grid(row=0, column=0, sticky='ew', **pad)
        hk_frame.columnconfigure(1, weight=1)
        self.hotkey_str = StringVar(value='ctrl+shift+m')
        ttk.Label(hk_frame, text='Toggle Hotkey (example: ctrl+shift+m)').grid(row=0, column=0, sticky='w')
        ttk.Entry(hk_frame, textvariable=self.hotkey_str).grid(row=0, column=1, sticky='ew')
        ttk.Button(hk_frame, text='Register Hotkey', command=self.register_hotkey).grid(row=0, column=2, sticky='e')
        global_frame = ttk.LabelFrame(tab_home, text='Global Timing')
        global_frame.grid(row=1, column=0, sticky='ew', **pad)
        global_frame.columnconfigure(1, weight=1)
        ttk.Label(global_frame, text='Cycle delay fixed (ms)').grid(row=0, column=0, sticky='w')
        self.global_fixed_ms = StringVar(value='100')
        ttk.Entry(global_frame, textvariable=self.global_fixed_ms, width=12).grid(row=0, column=1, sticky='w')
        ttk.Label(global_frame, text='Cycle jitter min,max (ms)').grid(row=0, column=2, sticky='w')
        self.global_jitter = StringVar(value='0,0')
        ttk.Entry(global_frame, textvariable=self.global_jitter, width=14).grid(row=0, column=3, sticky='w')
        chant_frame = ttk.LabelFrame(tab_home, text='Chant (combined sequence — steps separated by ;, parallel by ||)')
        chant_frame.grid(row=2, column=0, sticky='ew', **pad)
        chant_frame.columnconfigure(0, weight=1)
        self.chant_text = StringVar(value='')
        ttk.Entry(chant_frame, textvariable=self.chant_text).grid(row=0, column=0, padx=4, pady=4, sticky='ew')
        chant_help = (
            "Example: mouse|hold=14000|button=left || a|hold=14000 ; mouse|rel=180|dist=50|move=1000; mouse|hold=14000|button=left || a|hold=14000 ; stop"
        )
        ttk.Label(chant_frame, text=chant_help).grid(row=1, column=0, sticky='w')
        ctrl_frame = ttk.Frame(tab_home)
        ctrl_frame.grid(row=3, column=0, sticky='ew', **pad)
        ctrl_frame.columnconfigure(0, weight=1)
        self.toggle_label = ttk.Label(ctrl_frame, text='State: OFF')
        self.toggle_label.grid(row=0, column=0, sticky='w')
        ttk.Button(ctrl_frame, text='Start', command=self.gui_start).grid(row=0, column=1, sticky='e')
        ttk.Button(ctrl_frame, text='Stop', command=self.gui_stop).grid(row=0, column=2, sticky='e')
        ttk.Button(ctrl_frame, text='Quit', command=self.quit).grid(row=0, column=3, sticky='e')
        kb_frame = ttk.LabelFrame(tab_kb, text='Keyboard')
        kb_frame.grid(row=0, column=0, sticky='ew', **pad)
        kb_frame.columnconfigure(1, weight=1)
        self.enable_kb = IntVar(value=1)
        ttk.Checkbutton(kb_frame, text='Enable Keyboard', variable=self.enable_kb).grid(row=0, column=0, sticky='w')
        ttk.Label(kb_frame, text='Sequence (comma-separated)').grid(row=1, column=0, sticky='w')
        self.kb_sequence = StringVar(value='w+d,w,a,s|repeat=2')
        ttk.Entry(kb_frame, textvariable=self.kb_sequence).grid(row=1, column=1, columnspan=3, sticky='ew')
        ttk.Label(kb_frame, text='Mode').grid(row=2, column=0, sticky='w')
        self.kb_mode = StringVar(value='single')
        ttk.OptionMenu(kb_frame, self.kb_mode, 'single', 'single', 'hold', 'cps').grid(row=2, column=1, sticky='w')
        ttk.Label(kb_frame, text='Param (hold ms or cps range min,max)').grid(row=2, column=2, sticky='w')
        self.kb_param = StringVar(value='100')
        ttk.Entry(kb_frame, textvariable=self.kb_param, width=16).grid(row=2, column=3, sticky='w')
        ttk.Label(kb_frame, text='Per-action delay (fixed ms)').grid(row=3, column=0, sticky='w')
        self.kb_delay_fixed = StringVar(value='50')
        ttk.Entry(kb_frame, textvariable=self.kb_delay_fixed, width=12).grid(row=3, column=1, sticky='w')
        ttk.Label(kb_frame, text='Per-action jitter min,max (ms)').grid(row=3, column=2, sticky='w')
        self.kb_jitter = StringVar(value='0,20')
        ttk.Entry(kb_frame, textvariable=self.kb_jitter, width=14).grid(row=3, column=3, sticky='w')
        ttk.Label(kb_frame, text='Switch delay between pairs (ms)').grid(row=4, column=0, sticky='w')
        self.pair_switch_ms = StringVar(value='30')
        ttk.Entry(kb_frame, textvariable=self.pair_switch_ms, width=12).grid(row=4, column=1, sticky='w')
        mouse_frame = ttk.LabelFrame(tab_mouse, text='Mouse')
        mouse_frame.grid(row=0, column=0, sticky='ew', **pad)
        mouse_frame.columnconfigure(1, weight=1)
        self.enable_mouse = IntVar(value=1)
        ttk.Checkbutton(mouse_frame, text='Enable Mouse', variable=self.enable_mouse).grid(row=0, column=0, sticky='w')
        ttk.Label(mouse_frame, text='Mode').grid(row=1, column=0, sticky='w')
        self.mouse_mode = StringVar(value='single')
        ttk.OptionMenu(mouse_frame, self.mouse_mode, 'single', 'single', 'hold', 'cps', 'move').grid(row=1, column=1, sticky='w')
        ttk.Label(mouse_frame, text='Button').grid(row=1, column=2, sticky='w')
        self.mouse_button = StringVar(value='left')
        ttk.OptionMenu(mouse_frame, self.mouse_button, 'left', 'left', 'right').grid(row=1, column=3, sticky='w')
        ttk.Label(mouse_frame, text='Position (x,y) or blank').grid(row=2, column=0, sticky='w')
        self.mouse_pos = StringVar(value='')
        ttk.Entry(mouse_frame, textvariable=self.mouse_pos, width=16).grid(row=2, column=1, sticky='w')
        ttk.Label(mouse_frame, text='Offset jitter px').grid(row=2, column=2, sticky='w')
        self.mouse_jitter_px = StringVar(value='5')
        ttk.Entry(mouse_frame, textvariable=self.mouse_jitter_px, width=8).grid(row=2, column=3, sticky='w')
        ttk.Label(mouse_frame, text='Mouse param (hold ms or cps min,max)').grid(row=3, column=0, sticky='w')
        self.mouse_param = StringVar(value='100')
        ttk.Entry(mouse_frame, textvariable=self.mouse_param, width=16).grid(row=3, column=1, sticky='w')
        ttk.Label(mouse_frame, text='Per-action delay (fixed ms)').grid(row=4, column=0, sticky='w')
        self.mouse_delay_fixed = StringVar(value='50')
        ttk.Entry(mouse_frame, textvariable=self.mouse_delay_fixed, width=12).grid(row=4, column=1, sticky='w')
        ttk.Label(mouse_frame, text='Per-action jitter min,max (ms)').grid(row=4, column=2, sticky='w')
        self.mouse_jitter = StringVar(value='0,20')
        ttk.Entry(mouse_frame, textvariable=self.mouse_jitter, width=14).grid(row=4, column=3, sticky='w')
        move_frame = ttk.LabelFrame(mouse_frame, text='Movement / Ease Options')
        move_frame.grid(row=5, column=0, columnspan=4, sticky='ew', pady=(6,0))
        move_frame.columnconfigure(1, weight=1)
        ttk.Label(move_frame, text='Move style').grid(row=0, column=0, sticky='w')
        self.move_style = StringVar(value='ease-out')
        ttk.OptionMenu(move_frame, self.move_style, 'ease-out', 'linear', 'ease-out', 'ease-out+overshoot').grid(row=0, column=1, sticky='w')
        ttk.Label(move_frame, text='Duration (ms)').grid(row=0, column=2, sticky='w')
        self.move_dur = StringVar(value='200')
        ttk.Entry(move_frame, textvariable=self.move_dur, width=12).grid(row=0, column=3, sticky='w')
        ttk.Label(move_frame, text='Overshoot px min,max').grid(row=1, column=0, sticky='w')
        self.overshoot_px = StringVar(value='8,20')
        ttk.Entry(move_frame, textvariable=self.overshoot_px, width=16).grid(row=1, column=1, sticky='w')
        ttk.Label(move_frame, text='Axis-offset px min,max (applied on change)').grid(row=1, column=2, sticky='w')
        self.axis_offset_px = StringVar(value='0,6')
        ttk.Entry(move_frame, textvariable=self.axis_offset_px, width=16).grid(row=1, column=3, sticky='w')
        timing_frame = ttk.LabelFrame(tab_home, text='Inter-action Timing')
        timing_frame.grid(row=4, column=0, sticky='ew', **pad)
        timing_frame.columnconfigure(1, weight=1)
        ttk.Label(timing_frame, text='Inter-action fixed (ms)').grid(row=0, column=0, sticky='w')
        self.action_fixed = StringVar(value='20')
        ttk.Entry(timing_frame, textvariable=self.action_fixed, width=12).grid(row=0, column=1, sticky='w')
        ttk.Label(timing_frame, text='Inter-action jitter min,max (ms)').grid(row=0, column=2, sticky='w')
        self.action_jitter = StringVar(value='0,10')
        ttk.Entry(timing_frame, textvariable=self.action_jitter, width=14).grid(row=0, column=3, sticky='w')
        help_frame = ttk.LabelFrame(master, text='Sequence syntax examples')
        help_frame.grid(row=6, column=0, sticky='ew', **pad)
        help_frame.columnconfigure(0, weight=1)
        help_text = (
            "Chant example:\n"
            "mouse|hold=14000|button=left || a|hold=14000 ; mouse|rel=180|dist=50|move=1000; mouse|hold=14000|button=left || a|hold=14000 ; stop\n"
            "If Chant is non-empty it overrides separate KB/Mouse fields. This example works for Hypixel Skyblock farms, mining 2 rows at 175 speed."
        )
        ttk.Label(help_frame, text=help_text).grid(row=0, column=0, sticky='w')
        self.hotkey_listener = None
        self.register_hotkey()

    def hotkey_to_pynput(self, hk: str):
        parts = [p.strip().lower() for p in hk.split('+') if p.strip()]
        if not parts: return None
        pynput_parts = []
        for p in parts:
            if p in ('ctrl', 'control'): pynput_parts.append('<ctrl>')
            elif p in ('shift',): pynput_parts.append('<shift>')
            elif p in ('alt',): pynput_parts.append('<alt>')
            elif p in ('cmd', 'super', 'win'): pynput_parts.append('<cmd>')
            else: pynput_parts.append(p)
        return '+'.join(pynput_parts)

    def register_hotkey(self):
        hk = self.hotkey_str.get()
        if self.hotkey_listener:
            try: self.hotkey_listener.stop()
            except: pass
            self.hotkey_listener = None
        try:
            pynput_hk = self.hotkey_to_pynput(hk)
            if not pynput_hk:
                messagebox.showerror('Hotkey error', 'Invalid hotkey')
                return
            mapping = {pynput_hk: self.toggle_running}
            self.hotkey_listener = keyboard.GlobalHotKeys(mapping)
            self.hotkey_listener.start()
            messagebox.showinfo('Hotkey', f'Registered hotkey: {hk}')
        except Exception as e:
            messagebox.showerror('Hotkey error', f'Could not register hotkey: {e}')

    def gui_start(self):
        if self.controller.running:
            messagebox.showinfo('Already running', 'Automation already running')
            return
        self.toggle_running()

    def gui_stop(self):
        if not self.controller.running:
            messagebox.showinfo('Not running', 'Automation not running')
            return
        self.toggle_running()

    def toggle_running(self):
        if not self.controller.running:
            self.toggle_label.config(text='State: ON')
            job = self.automation_loop
            self.controller.start(job)
        else:
            self.controller.stop()
            self._cleanup_inputs()
            self.toggle_label.config(text='State: OFF')

    def quit(self):
        try:
            if self.hotkey_listener: self.hotkey_listener.stop()
        except: pass
        self.controller.stop()
        self._cleanup_inputs()
        self.master.quit()

    def automation_loop(self, stop_event: threading.Event):
        def parse_range(s, default=(0,0)):
            s = s.strip()
            if not s: return default
            parts = [p.strip() for p in s.split(',') if p.strip()]
            try:
                if len(parts) == 1:
                    v = int(parts[0]); return (v,v)
                elif len(parts) >= 2:
                    return (int(parts[0]), int(parts[1]))
            except:
                return default
        try:
            global_fixed = int(self.global_fixed_ms.get()) if self.global_fixed_ms.get().strip() else 0
        except:
            global_fixed = 0
        global_jitter = parse_range(self.global_jitter.get(), (0,0))
        try:
            action_fixed = int(self.action_fixed.get()) if self.action_fixed.get().strip() else 0
        except:
            action_fixed = 0
        action_jitter = parse_range(self.action_jitter.get(), (0,0))
        try:
            move_dur = int(self.move_dur.get())
        except:
            move_dur = 200
        overshoot_range = parse_range(self.overshoot_px.get(), (0,0))
        axis_offset_range = parse_range(self.axis_offset_px.get(), (0,0))
        move_style = self.move_style.get()
        chant_raw = self.chant_text.get().strip()
        if chant_raw:
            chant_steps = parse_chant(chant_raw)
            while not stop_event.is_set():
                for step in chant_steps:
                    if stop_event.is_set(): break
                    stop_found = any(act.get('device') == 'stop' for act in step)
                    if stop_found:
                        stop_event.set()
                        self._cleanup_inputs()
                        return
                    threads = []
                    for act in step:
                        if act.get('device') == 'kb':
                            t = threading.Thread(target=self._kb_action_once, args=(act, stop_event), daemon=True)
                            threads.append(t); t.start()
                        elif act.get('device') == 'mouse':
                            t = threading.Thread(target=self._mouse_action_from_chant, args=(act, move_style, move_dur, overshoot_range, axis_offset_range, stop_event), daemon=True)
                            threads.append(t); t.start()
                    for t in threads:
                        while t.is_alive():
                            if stop_event.is_set(): break
                            t.join(timeout=0.05)
                    total_global = global_fixed + random.randint(*global_jitter)
                    self._sleep_ms(total_global, stop_event)
            self._cleanup_inputs()
            return
        while not stop_event.is_set():
            enable_kb = bool(self.enable_kb.get())
            enable_mouse = bool(self.enable_mouse.get())
            kb_seq_raw = self.kb_sequence.get()
            kb_actions = parse_sequence(kb_seq_raw)
            mmode = self.mouse_mode.get()
            btn = self.mouse_button.get()
            btn_obj = MouseButton.left if btn == 'left' else MouseButton.right
            mouse_pos_raw = self.mouse_pos.get().strip()
            def parse_mouse_pos(raw):
                if not raw:
                    return None
                try:
                    x_str, y_str = [s.strip() for s in raw.split(',')]
                    return (int(x_str), int(y_str))
                except:
                    return None
            mouse_pos = parse_mouse_pos(mouse_pos_raw)
            need_move = False
            if mouse_pos and mmode in ('single','hold','cps','move'):
                need_move = True
            move_thread = None
            if enable_mouse and need_move:
                dur_ms = move_dur
                move_thread = threading.Thread(
                    target=self._mouse_move_to,
                    args=(mouse_pos, int(self.mouse_jitter_px.get() or 0), dur_ms, move_style, overshoot_range, axis_offset_range, stop_event),
                    daemon=True
                )
                move_thread.start()
            threads = []
            if enable_kb and kb_actions:
                t_kb = threading.Thread(target=self._kb_worker, args=(stop_event, kb_actions, self.kb_mode.get(), self._parse_range_from_string(self.kb_param.get()), int(self.kb_delay_fixed.get() or 0), self._parse_range_from_string(self.kb_jitter.get()), int(self.pair_switch_ms.get() or 0), int(action_fixed or 0), self._parse_range_from_string(self.action_jitter.get())), daemon=True)
                threads.append(t_kb); t_kb.start()
            if enable_mouse:
                t_mouse = threading.Thread(target=self._mouse_worker, args=(stop_event, mmode, mouse_pos, int(self.mouse_jitter_px.get() or 0), self._parse_range_from_string(self.mouse_param.get()), int(self.mouse_delay_fixed.get() or 0), self._parse_range_from_string(self.mouse_jitter.get()), btn_obj, int(self.pair_switch_ms.get() or 0), int(action_fixed or 0), self._parse_range_from_string(self.action_jitter.get())), daemon=True)
                threads.append(t_mouse); t_mouse.start()
            for t in threads:
                while t.is_alive():
                    if stop_event.is_set(): break
                    t.join(timeout=0.05)
            if move_thread:
                while move_thread.is_alive():
                    if stop_event.is_set(): break
                    move_thread.join(timeout=0.05)
            total_global = global_fixed + random.randint(*global_jitter)
            self._sleep_ms(total_global, stop_event)
        self._cleanup_inputs()

    def _parse_range_from_string(self, s):
        s = s.strip()
        if not s:
            return (0,0)
        parts = [p.strip() for p in s.split(',') if p.strip()]
        try:
            if len(parts) == 1:
                v = int(parts[0]); return (v,v)
            else:
                return (int(parts[0]), int(parts[1]))
        except:
            return (0,0)

    def _kb_action_once(self, act, stop_event=None):
        se = stop_event or getattr(self.controller, '_stop_event', None)
        keys = act.get('keys', [])
        simul = act.get('simul', False)
        hold = act.get('hold', None)
        repeat = act.get('repeat', 1)
        try:
            for _ in range(repeat):
                if se and se.is_set(): break
                if simul:
                    for k in keys:
                        if k is None: continue
                        try: self.controller.kc.press(k); self._pressed_keys.add(k)
                        except: pass
                    if hold:
                        waited=0.0; tgt=hold/1000.0
                        while waited < tgt:
                            if se and se.is_set(): break
                            time.sleep(min(0.02, tgt - waited)); waited += min(0.02, tgt - waited)
                    else:
                        time.sleep(0.01)
                    for k in reversed(keys):
                        try: self.controller.kc.release(k)
                        except: pass
                        if k in self._pressed_keys: self._pressed_keys.discard(k)
                else:
                    for k in keys:
                        if se and se.is_set(): break
                        if k is None: continue
                        if hold:
                            try: self.controller.kc.press(k); self._pressed_keys.add(k)
                            except: pass
                            waited=0.0; tgt=hold/1000.0
                            while waited < tgt:
                                if se and se.is_set(): break
                                time.sleep(min(0.02, tgt-waited)); waited += min(0.02, tgt-waited)
                            try: self.controller.kc.release(k)
                            except: pass
                            if k in self._pressed_keys: self._pressed_keys.discard(k)
                        else:
                            try: self.controller.kc.press(k); self._pressed_keys.add(k)
                            except: pass
                            time.sleep(0.01)
                            try: self.controller.kc.release(k)
                            except: pass
                            if k in self._pressed_keys: self._pressed_keys.discard(k)
        except Exception as e:
            print('_kb_action_once error:', e)
        finally:
            if se and se.is_set():
                self._cleanup_inputs()

    def _mouse_action_from_chant(self, act, move_style, default_move_dur, overshoot_range, axis_offset_range, stop_event=None):
        se = stop_event or getattr(self.controller, '_stop_event', None)
        for _ in range(act.get('repeat', 1)):
            if se and se.is_set(): break

            pos = act.get('pos', None)
            rel = act.get('rel', None)
            dist = act.get('dist', None) or 400

            if rel is not None:
                try:
                    angle = float(rel)
                except Exception:
                    angle = 0.0
                cur = self.controller.mc.position
                rad = math.radians(angle)
                dx = math.cos(rad) * dist
                dy = math.sin(rad) * dist

                eps = 0.5
                tx = round(cur[0] + dx)
                ty = round(cur[1] + dy)
                if abs(dy) < eps:
                    ty = cur[1]
                if abs(dx) < eps:
                    tx = cur[0]

                pos = (int(tx), int(ty))

            move_ms = act.get('move', None) or default_move_dur

            mv_thread = None
            if pos is not None:
                mv_thread = threading.Thread(
                    target=self._mouse_move_to,
                    args=(pos, int(self.mouse_jitter_px.get() or 0), int(move_ms), move_style, overshoot_range, axis_offset_range, se),
                    daemon=True
                )
                mv_thread.start()

            hold_ms = act.get('hold', None)
            btn = MouseButton.left if act.get('button','left') == 'left' else MouseButton.right

            if hold_ms:
                waited = 0.0
                wait_target = min(0.15, move_ms / 1000.0)
                while waited < wait_target:
                    if se and se.is_set(): break
                    time.sleep(0.01); waited += 0.01

                try:
                    self.controller.mc.press(btn)
                    self._pressed_buttons.add(btn)
                except Exception:
                    pass

                start = time.time(); target = hold_ms / 1000.0
                while (time.time() - start) < target:
                    if se and se.is_set(): break
                    time.sleep(0.02)
                try:
                    self.controller.mc.release(btn)
                except:
                    pass
                if btn in self._pressed_buttons:
                    self._pressed_buttons.discard(btn)
            else:
                if mv_thread:
                    waited = 0.0; tgt = min(0.2, move_ms / 1000.0)
                    while waited < tgt:
                        if se and se.is_set(): break
                        time.sleep(0.01); waited += 0.01

                try:
                    if pos is not None:
                        self.controller.mc.position = (int(pos[0]), int(pos[1]))
                    self.controller.mc.click(btn)
                except Exception:
                    pass

            if mv_thread:
                while mv_thread.is_alive():
                    if se and se.is_set(): break
                    mv_thread.join(timeout=0.05)

            if se and se.is_set(): break

    def _mouse_move_semicircle(self, start, target, dur_ms, style, stop_event=None, clockwise=True):
        try:
            se = stop_event or getattr(self.controller, '_stop_event', None)
            sx, sy = start
            tx, ty = target
            cx = (sx + tx) / 2.0
            cy = (sy + ty) / 2.0
            dx = sx - cx
            dy = sy - cy
            r = math.hypot(dx, dy)
            if r < 1e-6:
                self.controller.mc.position = (int(tx), int(ty))
                return
            start_ang = math.atan2(sy - cy, sx - cx)
            if clockwise:
                end_ang = start_ang - math.pi
            else:
                end_ang = start_ang + math.pi

            total_steps = max(1, int(max(1, dur_ms) / 8))
            for i in range(1, total_steps + 1):
                if se and se.is_set(): break
                t = i / total_steps
                if style == 'linear':
                    ang = start_ang + (end_ang - start_ang) * t
                else:
                    e = ease_out_cubic(t)
                    ang = start_ang + (end_ang - start_ang) * e
                nx = cx + math.cos(ang) * r
                ny = cy + math.sin(ang) * r
                self.controller.mc.position = (int(nx), int(ny))
                self._sleep_ms(dur_ms / total_steps * 1.0, stop_event)
            try:
                self.controller.mc.position = (int(tx), int(ty))
            except Exception:
                pass
        except Exception as e:
            print('Mouse semicircle move error:', e)

    def _kb_worker(self, stop_event, kb_actions, kb_mode, kb_param_range, kb_delay_fixed, kb_jitter, pair_switch, action_fixed, action_jitter):
        try:
            if kb_mode == 'single':
                for idx, act in enumerate(kb_actions):
                    if stop_event.is_set(): break
                    self._do_kb_action(act, kb_delay_fixed, kb_jitter, stop_event)
                    self._sleep_ms(pair_switch, stop_event)
                    self._sleep_ms(action_fixed + random.randint(*action_jitter), stop_event)
            elif kb_mode == 'hold':
                hold_ms = kb_param_range[0]
                for idx, act in enumerate(kb_actions):
                    if stop_event.is_set(): break
                    h = act['hold'] if act['hold'] is not None else hold_ms
                    act_copy = dict(act); act_copy['hold'] = h
                    self._do_kb_action(act_copy, kb_delay_fixed, kb_jitter, stop_event)
                    self._sleep_ms(pair_switch, stop_event)
                    self._sleep_ms(action_fixed + random.randint(*action_jitter), stop_event)
            elif kb_mode == 'cps':
                min_ms, max_ms = kb_param_range
                for idx, act in enumerate(kb_actions):
                    if stop_event.is_set(): break
                    self._do_kb_action(act, kb_delay_fixed, kb_jitter, stop_event)
                    self._sleep_ms(random.randint(min_ms, max_ms), stop_event)
                    self._sleep_ms(pair_switch, stop_event)
                    self._sleep_ms(action_fixed + random.randint(*action_jitter), stop_event)
        except Exception as e:
            print('KB worker error:', e)
        finally:
            if stop_event.is_set():
                self._cleanup_inputs()

    def _mouse_worker(self, stop_event, mmode, mouse_pos, mouse_jitter_px, mouse_param_range, mouse_delay_fixed, mouse_jitter, btn_obj, pair_switch, action_fixed, action_jitter):
        try:
            if stop_event.is_set(): return
            if mmode == 'single':
                self._mouse_click_at(mouse_pos, mouse_jitter_px, btn_obj)
                self._sleep_ms(mouse_delay_fixed + random.randint(*mouse_jitter), stop_event)
                self._sleep_ms(action_fixed + random.randint(*action_jitter), stop_event)
            elif mmode == 'hold':
                hold_ms = mouse_param_range[0]
                self._mouse_hold(hold_ms, mouse_pos, mouse_jitter_px, btn_obj, stop_event)
                self._sleep_ms(action_fixed + random.randint(*action_jitter), stop_event)
            elif mmode == 'cps':
                min_ms, max_ms = mouse_param_range
                self._mouse_click_at(mouse_pos, mouse_jitter_px, btn_obj)
                self._sleep_ms(random.randint(min_ms, max_ms), stop_event)
                self._sleep_ms(action_fixed + random.randint(*action_jitter), stop_event)
            elif mmode == 'move':
                pass
        except Exception as e:
            print('Mouse worker error:', e)
        finally:
            if stop_event.is_set():
                self._cleanup_inputs()

    def _do_kb_action(self, act, delay_fixed, delay_jitter, stop_event=None):
        se = stop_event or getattr(self.controller, '_stop_event', None)
        keys = act.get('keys', [])
        simul = act.get('simul', False)
        hold = act.get('hold', None)
        repeat = act.get('repeat', 1)
        for _ in range(repeat):
            if se and se.is_set(): break
            if simul:
                try:
                    for k in keys:
                        if k is None: continue
                        try: self.controller.kc.press(k); self._pressed_keys.add(k)
                        except: pass
                    if hold:
                        waited=0.0; target=hold/1000.0
                        while waited < target:
                            if se and se.is_set(): break
                            sleep_chunk = min(0.02, target - waited)
                            time.sleep(sleep_chunk); waited += sleep_chunk
                    else:
                        waited=0.0; target=0.01
                        while waited < target:
                            if se and se.is_set(): break
                            time.sleep(0.01); waited += 0.01
                except Exception as e:
                    print('Key combo press error:', e)
                finally:
                    try:
                        for k in reversed(keys):
                            if k is None: continue
                            try: self.controller.kc.release(k)
                            except: pass
                            if k in self._pressed_keys: self._pressed_keys.discard(k)
                    except Exception as e:
                        print('Key combo release error:', e)
            else:
                for k in keys:
                    if se and se.is_set(): break
                    if k is None: continue
                    try:
                        if hold:
                            try: self.controller.kc.press(k); self._pressed_keys.add(k)
                            except: pass
                            waited=0.0; target=hold/1000.0
                            while waited < target:
                                if se and se.is_set(): break
                                sleep_chunk = min(0.02, target - waited)
                                time.sleep(sleep_chunk); waited += sleep_chunk
                            try: self.controller.kc.release(k)
                            except: pass
                            if k in self._pressed_keys: self._pressed_keys.discard(k)
                        else:
                            try: self.controller.kc.press(k); self._pressed_keys.add(k)
                            except: pass
                            waited=0.0; target=0.01
                            while waited < target:
                                if se and se.is_set(): break
                                time.sleep(0.01); waited += 0.01
                            try: self.controller.kc.release(k)
                            except: pass
                            if k in self._pressed_keys: self._pressed_keys.discard(k)
                    except Exception as e:
                        print('Key press error:', e)
            self._sleep_ms(delay_fixed + random.randint(*delay_jitter), stop_event)

    def _mouse_click_at(self, pos, jitter_px, btn_obj):
        try:
            if pos:
                tx, ty = pos
            else:
                cur = self.controller.mc.position
                tx, ty = cur
            dx = random.randint(-jitter_px, jitter_px) if jitter_px else 0
            dy = random.randint(-jitter_px, jitter_px) if jitter_px else 0
            tx += dx; ty += dy
            self.controller.mc.position = (tx, ty)
            time.sleep(0.005)
            self.controller.mc.click(btn_obj)
        except Exception as e:
            print('Mouse click error:', e)

    def _mouse_hold(self, ms, pos, jitter_px, btn_obj, stop_event=None):
        try:
            se = stop_event or getattr(self.controller, '_stop_event', None)
            if pos:
                tx, ty = pos
            else:
                cur = self.controller.mc.position
                tx, ty = cur
            dx = random.randint(-jitter_px, jitter_px) if jitter_px else 0
            dy = random.randint(-jitter_px, jitter_px) if jitter_px else 0
            tx += dx; ty += dy
            self.controller.mc.position = (tx, ty)
            try:
                self.controller.mc.press(btn_obj)
                self._pressed_buttons.add(btn_obj)
            except Exception:
                pass
            start = time.time()
            target = ms / 1000.0
            while (time.time() - start) < target:
                if se and se.is_set(): break
                time.sleep(0.01)
            try:
                self.controller.mc.release(btn_obj)
            except Exception:
                pass
            if btn_obj in self._pressed_buttons:
                self._pressed_buttons.discard(btn_obj)
        except Exception as e:
            print('Mouse hold error:', e)

    def _mouse_move_to(self, pos, jitter_px, dur_ms, style, overshoot_range, axis_offset_range, stop_event=None):
        try:
            se = stop_event or getattr(self.controller, '_stop_event', None)
            tx, ty = pos
            tx += random.randint(-jitter_px, jitter_px) if jitter_px else 0
            ty += random.randint(-jitter_px, jitter_px) if jitter_px else 0

            if hasattr(self, '_last_mouse_target') and self._last_mouse_target is not None:
                lx, ly = self._last_mouse_target
                if (lx, ly) != (tx, ty):
                    ax_min, ax_max = axis_offset_range
                    if ax_max >= ax_min and ax_max > 0:
                        if tx != lx:
                            off_x = random.randint(ax_min, ax_max)
                            tx += random.choice((-1, 1)) * off_x
                        if ty != ly:
                            off_y = random.randint(ax_min, ax_max)
                            ty += random.choice((-1, 1)) * off_y

            start = self.controller.mc.position
            sx = float(start[0]); sy = float(start[1])
            txf = float(tx); tyf = float(ty)

            dx = txf - sx
            dy = tyf - sy

            eps = 0.5
            move_x = abs(dx) >= eps
            move_y = abs(dy) >= eps

            total_steps = max(1, int(max(1, dur_ms) / 8))

            def step_set(nx, ny):
                try:
                    self.controller.mc.position = (int(round(nx)), int(round(ny)))
                except Exception:
                    pass

            if style == 'linear':
                for i in range(1, total_steps + 1):
                    if se and se.is_set(): break
                    t = i / total_steps
                    nx = sx + dx * t if move_x else sx
                    ny = sy + dy * t if move_y else sy
                    step_set(nx, ny)
                    self._sleep_ms(dur_ms / total_steps, stop_event)
            elif style == 'ease-out':
                for i in range(1, total_steps + 1):
                    if se and se.is_set(): break
                    t = i / total_steps; e = ease_out_cubic(t)
                    nx = sx + dx * e if move_x else sx
                    ny = sy + dy * e if move_y else sy
                    step_set(nx, ny)
                    self._sleep_ms(dur_ms / total_steps, stop_event)
            elif style == 'ease-out+overshoot':
                dist = math.hypot(dx, dy)
                os_min, os_max = overshoot_range
                overshoot_px = random.randint(os_min, os_max) if os_max >= os_min and os_max > 0 else 0

                if dist <= 1e-6:
                    try:
                        self.controller.mc.position = (int(round(txf)), int(round(tyf)))
                    except Exception:
                        pass
                    self._last_mouse_target = (int(txf), int(tyf))
                    return

                if not move_x and move_y:
                    ux, uy = 0.0, math.copysign(1.0, dy)
                elif not move_y and move_x:
                    ux, uy = math.copysign(1.0, dx), 0.0
                else:
                    ux, uy = dx / dist, dy / dist

                ox = ux * overshoot_px
                oy = uy * overshoot_px
                ox_t = txf + ox
                oy_t = tyf + oy

                phase1_steps = max(1, int(total_steps * 0.7))
                phase2_steps = max(1, total_steps - phase1_steps)

                for i in range(1, phase1_steps + 1):
                    if se and se.is_set(): break
                    t = i / phase1_steps; e = ease_out_cubic(t)
                    nx = sx + (ox_t - sx) * e if move_x or move_y else sx
                    ny = sy + (oy_t - sy) * e if move_y or move_x else sy
                    if not move_x: nx = sx
                    if not move_y: ny = sy
                    step_set(nx, ny)
                    self._sleep_ms(dur_ms / total_steps, stop_event)

                for i in range(1, phase2_steps + 1):
                    if se and se.is_set(): break
                    t = i / phase2_steps; e = ease_in_out_cubic(t)
                    nx = ox_t + (txf - ox_t) * e if move_x else txf
                    ny = oy_t + (tyf - oy_t) * e if move_y else tyf
                    if not move_x: nx = txf
                    if not move_y: ny = tyf
                    step_set(nx, ny)
                    self._sleep_ms(dur_ms / total_steps, stop_event)
            else:
                for i in range(1, total_steps + 1):
                    if se and se.is_set(): break
                    t = i / total_steps
                    nx = sx + dx * t if move_x else sx
                    ny = sy + dy * t if move_y else sy
                    step_set(nx, ny)
                    self._sleep_ms(dur_ms / total_steps, stop_event)

            try:
                self.controller.mc.position = (int(round(txf)), int(round(tyf)))
            except Exception:
                pass

            self._last_mouse_target = (int(round(txf)), int(round(tyf)))
        except Exception as e:
            print('Mouse move error:', e)

    def _sleep_ms(self, ms, stop_event=None):
        if ms <= 0: return
        se = stop_event or getattr(self.controller, '_stop_event', None)
        total = ms_to_sec(ms)
        slept = 0.0
        while slept < total:
            if se and se.is_set(): return
            chunk = min(0.02, total - slept)
            time.sleep(chunk); slept += chunk

    def _cleanup_inputs(self):
        for k in list(self._pressed_keys):
            try: self.controller.kc.release(k)
            except: pass
            self._pressed_keys.discard(k)
        for b in list(self._pressed_buttons):
            try: self.controller.mc.release(b)
            except: pass
            self._pressed_buttons.discard(b)

if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        pass
    finally:
        app.quit()
