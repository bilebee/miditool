import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import rtmidi
from mido import MidiFile, MidiTrack, Message, MetaMessage
import os
import queue
from datetime import datetime

class MidiCoreThread(threading.Thread):
    """优化后的核心线程：专注事件收集与转发"""
    def __init__(self):
        super().__init__()
        self.input = rtmidi.MidiIn()
        self.output = rtmidi.MidiOut()
        self.recording = False
        self.virtual_output = False
        self._event_queue = queue.Queue()
        self._control_queue = queue.Queue()
        self.input.set_callback(self._midi_callback)
        self.bpm = 120  # 新增BPM属性
        # 移除锁机制
        self.events = []
        # 修改默认保存路径为桌面
        self.save_path = os.path.join(os.path.expanduser("~"), "Desktop")
        self.filename = "recording.mid"
        self._active = True  # 新增线程活动状态标识

    def _midi_callback(self, event, _):
        message, timestamp = event
        if self.recording:
            # 直接操作列表（单线程无需锁）
            self.events.append({
                'data': bytes(message),
                'timestamp': time.time()
            })
        if self.virtual_output:
            return
        self.output.send_message(message)

    def run(self):
        """简化的主循环"""
        while self._active:
            self._process_control()
            time.sleep(0.001)
        self._close_ports()

    def _process_control(self):
        try:
            cmd, *args = self._control_queue.get_nowait()
            if cmd == 'CONNECT':
                self._connect_device(*args)
            elif cmd == 'START_RECORD':
                self._start_recording(*args)
            elif cmd == 'STOP_RECORD':
                self._stop_recording()
                self._active = False  # 先保存再停止线程
        except queue.Empty:
            pass

    def _connect_device(self, in_port, out_port):
        try:
            self._close_ports()
            if in_port < self.input.get_port_count():
                self.input.open_port(in_port)
            self.virtual_output = (out_port == -1)
            if not self.virtual_output and out_port < self.output.get_port_count():
                self.output.open_port(out_port)
            self._event_queue.put(('STATUS', '设备已连接'))
        except rtmidi.RtMidiError as e:
            self._event_queue.put(('ERROR', str(e)))

    def _start_recording(self, save_path, filename, bpm):
        self.events.clear()
        self.save_path = save_path
        self.filename = filename
        self.bpm = bpm  # 设置BPM
        self.recording = True

    def _stop_recording(self):
        """停止录制时立即保存并终止线程"""
        self.recording = False
        self._save_midi()

    def _save_midi(self):
        if not self.events:
           return

        mid = MidiFile()
        track = MidiTrack()
        mid.tracks.append(track)

        # 添加BPM元事件（放在第一个位置）
        microseconds_per_beat = int(60000000 / self.bpm)
        track.append(MetaMessage('set_tempo', tempo=microseconds_per_beat, time=0))

        # 查找第一个note_on事件作为基准时间
        base_time = None
        first_note_index = 0
        for i, event in enumerate(self.events):
            try:
                msg = Message.from_bytes(event['data'])
                if msg.type == 'note_on':
                    base_time = event['timestamp']
                    first_note_index = i
                    break
            except:
                continue

        # 如果没有note_on事件，使用第一个事件作为基准
        if base_time is None:
            base_time = self.events[0]['timestamp']
            first_note_index = 0

        prev_time = base_time

        # 处理所有事件（包含第一个note_on之前的其他事件）
        for event in self.events:
            current_time = event['timestamp']
            delta = int((current_time - prev_time) * 1000)  # 转换为毫秒
        
            try:
                msg = Message.from_bytes(event['data'])
            
                # 时间校正逻辑
                if delta < 0:
                    if msg.type == 'note_off':
                        continue  # 直接跳过note_off事件
                    else:
                        delta = 0  # 其他类型设置为0
            
                # 特殊处理note_on的0时间（可选）
                if delta == 0 and msg.type == 'note_on':
                    track.append(msg.copy(time=0))
                else:
                    track.append(msg.copy(time=delta))
                
                prev_time = current_time  # 仅更新有效事件的时间戳
            except Exception as e:
                self._event_queue.put(('ERROR', f"无效消息：{str(e)}"))

        full_path = os.path.join(self.save_path, self.filename)
        try:
            mid.save(full_path)
            self._event_queue.put(('STATUS', f'文件已保存：{full_path}'))
        except Exception as e:
            self._event_queue.put(('ERROR', f'保存失败：{str(e)}'))

    def _close_ports(self):
        if self.input.is_port_open():
            self.input.close_port()
        if self.output.is_port_open():
            self.output.close_port()

    @staticmethod
    def parse_midi_log(log_content, bpm=120):
        """完全基于日志元数据的解析方法"""
        # 初始化默认值
        tpb = 120  # 默认TPB
        tempo = 500000  # 默认tempo (120 BPM)
        time_sig = (4, 4, 24, 8)  # 默认拍号

        # 预扫描日志获取元数据
        for line in log_content.split('\n'):
            line = line.strip()
            if not line or line.startswith('#'):  # 跳过注释
                continue
            
            parts = line.split()
            if len(parts) < 2:  # 基本长度校验
                continue
            
            # 仅处理元数据行 ---------------------------------------------------
            if parts[0] == 'MFile':
                if len(parts) >= 4:
                    try:
                        tpb = int(parts[3])
                    except ValueError:
                        pass
            elif parts[1] == 'TimeSig':
                if len(parts) >= 5 and '/' in parts[2]:
                    try:
                        numerator, denominator = map(int, parts[2].split('/'))
                        time_sig = (
                            numerator,
                            denominator,
                            int(parts[3]),
                            int(parts[4])
                        )
                    except (ValueError, IndexError):
                        pass
            elif parts[1] == 'Tempo':
                if len(parts) >= 3:
                    try:
                        tempo = int(parts[2])
                    except ValueError:
                        pass
        # ----------------------------------------------------------------

        # 使用日志中的实际参数创建MIDI
        mid = MidiFile(ticks_per_beat=tpb)
        track = MidiTrack()
        mid.tracks.append(track)
    
        # 添加初始元事件
        track.append(MetaMessage('time_signature', 
            numerator=time_sig[0],
            denominator=time_sig[1],
            clocks_per_click=time_sig[2],
            notated_32nd_notes_per_beat=time_sig[3],
            time=0))
        track.append(MetaMessage('set_tempo', tempo=tempo, time=0))
    
        prev_ticks = 0
        for i, line in enumerate(log_content.split('\n')):
            line = line.strip()
            if not line or line.startswith('#'):  # 跳过空行和注释
                continue
            
            parts = line.split()
            if len(parts) < 2:  # 强化校验：至少包含时间戳和事件类型
                continue
            
            if parts[0] in ['MFile', 'MTrk', 'TrkEnd']:
                continue  # 跳过文件头标记
            
            # 新增事件类型校验
            event_type = parts[1].lower()
            supported_events = ['timesig', 'tempo', 'on', 'off', 'par', 'prch', 'pb']
            if event_type not in supported_events:
                continue

            try:
                current_ticks = int(parts[0])
                delta = current_ticks - prev_ticks
                prev_ticks = current_ticks
                
                # 统一参数解析
                params = {}
                for p in parts[2:]:
                    if '=' in p:
                        key, val = p.split('=', 1)
                        params[key] = int(val)
                    elif p.isdigit():  # 处理类似 Pb 8192 这样的数值
                        params['value'] = int(p)

                channel = params.get('ch', 1) - 1  # 通道号转换

                # 处理不同事件类型
                msg = None
                if event_type == 'par' and 'c' in params:
                    msg = Message('control_change',
                            channel=channel,
                            control=params['c'],
                            value=params['v'])
                            
                elif event_type == 'prch' and 'p' in params:
                    msg = Message('program_change',
                            channel=channel,
                            program=params['p'])
                            
                elif event_type == 'pb' and 'v' in params:
                    # 转换14位Pitch Bend值（假设v是0-16383）
                    value = min(max(params['v'], 0), 16383)
                    msg = Message('pitchwheel',
                            channel=channel,
                            pitch=value - 8192)  # 转换为-8192到8191
                            
                elif event_type in ['on', 'note_on']:
                    if 'n' in params:
                        msg = Message('note_on',
                                channel=channel,
                                note=params['n'],
                                velocity=params.get('v', 64))
                elif event_type in ['off', 'note_off']:
                    if 'n' in params:
                        msg = Message('note_off',
                                channel=channel,
                                note=params['n'],
                                velocity=params.get('v', 64))

                if msg:
                    track.append(msg.copy(time=delta))
                
            except Exception as e:
                raise ValueError(f"第{i+1}行解析失败：{str(e)}")
            
        return mid

class MidiRecorderApp:
    def __init__(self, master):
        self.master = master
        self.core_thread = MidiCoreThread()
        self.core_thread.start()
                # 新增分辨率变量初始化
        self.resolution_var = tk.IntVar(value=480)  # 初始化分辨率变量
        self._setup_gui()
        self._setup_event_handling()
        self._refresh_devices()

        
    def _setup_gui(self):
        self.master.title("MIDI Recorder v3.0")
        self.master.geometry("800x600")
        
        main_frame = ttk.Frame(self.master)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 设备选择区域
        device_frame = ttk.LabelFrame(main_frame, text="设备设置")
        device_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(device_frame, text="输入设备:").grid(row=0, column=0, padx=5, sticky=tk.W)
        self.input_combo = ttk.Combobox(device_frame, state="readonly", width=30)
        self.input_combo.grid(row=0, column=1, padx=5)
        
        ttk.Label(device_frame, text="输出设备:").grid(row=0, column=2, padx=5, sticky=tk.W)
        self.output_combo = ttk.Combobox(device_frame, state="readonly", width=30)
        self.output_combo.grid(row=0, column=3, padx=5)
        
        ttk.Button(device_frame, text="刷新设备", command=self._refresh_devices).grid(row=0, column=4, padx=5)

        # 文件保存区域
        save_frame = ttk.LabelFrame(main_frame, text="文件设置")
        save_frame.pack(fill=tk.X, pady=5)
        ttk.Label(save_frame, text="BPM:").grid(row=1, column=2, padx=5, sticky=tk.E)
        self.bpm_var = tk.IntVar(value=120)
        ttk.Spinbox(save_frame, from_=20, to=300, textvariable=self.bpm_var, width=5).grid(row=1, column=3, padx=5)
        ttk.Label(save_frame, text="保存路径:").grid(row=0, column=0, padx=5, sticky=tk.W)
        # 修改默认路径显示为桌面
        self.path_var = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        ttk.Entry(save_frame, textvariable=self.path_var, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(save_frame, text="浏览", command=self._choose_path).grid(row=0, column=2, padx=5)
        
        ttk.Label(save_frame, text="文件名:").grid(row=1, column=0, padx=5, sticky=tk.W)
        self.name_var = tk.StringVar(value=f"recording_{datetime.now().strftime('%Y%m%d_%H%M%S')}.mid")
        ttk.Entry(save_frame, textvariable=self.name_var).grid(row=1, column=1, padx=5, sticky=tk.W)

        # 分辨率设置
        ttk.Label(save_frame, text="分辨率:").grid(row=1, column=4, padx=5, sticky=tk.E)
        ttk.Spinbox(save_frame, from_=96, to=960, increment=48, 
                   textvariable=self.resolution_var, width=6).grid(row=1, column=5, padx=5)

        # 控制按钮
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(pady=10)
        self.record_btn = ttk.Button(control_frame, text="开始录制", command=self._toggle_recording)
        self.record_btn.pack(side=tk.LEFT, padx=5)

        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="日志")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log = scrolledtext.ScrolledText(log_frame)
        self.log.pack(fill=tk.BOTH, expand=True)

        # 转换功能区
        convert_frame = ttk.LabelFrame(main_frame, text="日志转换")
        convert_frame.pack(fill=tk.X, pady=5)

        ttk.Button(convert_frame, 
                  text="导入MIDI日志", 
                  command=self._import_log).pack(side=tk.LEFT, padx=5)

    def _setup_event_handling(self):
        self.master.after(100, self._process_core_events)

    def _refresh_devices(self):
        devices = MidiRecorderApp.list_devices()  # 修改为调用静态方法
        self.input_combo['values'] = [name for _, name in devices['inputs']]
        self.output_combo['values'] = [name for _, name in devices['outputs']]
        self.input_combo.current(0) if devices['inputs'] else None
        self.output_combo.current(0)

    def _choose_path(self):
        path = filedialog.askdirectory()
        if path:
            self.path_var.set(path)

    def _toggle_recording(self):
        """每次录制创建新线程实例"""
        if not self.core_thread.is_alive():
            # 创建新线程实例
            self.core_thread = MidiCoreThread()
            self.core_thread.start()
            
            # 发送设备连接命令
            in_idx = self.input_combo.current()
            out_idx = self.output_combo.current()
            devices = MidiRecorderApp.list_devices()  # 修改为调用静态方法
            self.core_thread._control_queue.put((  # 修改为调用静态方法
                'CONNECT',
                devices['inputs'][in_idx][0],
                devices['outputs'][out_idx][0]
            ))
            
            # 发送开始录制命令
            self.core_thread._control_queue.put(((
                'START_RECORD',
                self.path_var.get(),
                self.name_var.get(),
                self.bpm_var.get()  # 新增BPM参数
            )))
            self.record_btn.config(text="停止录制")
        else:
            # 发送终止命令（会自动关闭线程）
            self.core_thread._control_queue.put(('STOP_RECORD',))
            self.record_btn.config(text="开始录制")

    def _process_core_events(self):
        while True:
            try:
                event_type, data = self.core_thread._event_queue.get_nowait()
                if event_type == 'ERROR':
                    self._log_message(data, is_error=True)
                elif event_type == 'STATUS':
                    self._log_message(data)
            except queue.Empty:
                break
        self.master.after(100, self._process_core_events)

    def _log_message(self, message, is_error=False):
        tag = "ERROR" if is_error else "INFO"
        self.log.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {tag} - {message}\n")
        self.log.tag_config("ERROR", foreground="red")
        self.log.see(tk.END)
        if is_error:
            self.record_btn.config(text="开始录制")

    @staticmethod
    def list_devices():
        """静态方法获取设备列表"""
        input = rtmidi.MidiIn()
        output = rtmidi.MidiOut()
        inputs = [(i, input.get_port_name(i)) for i in range(input.get_port_count())]
        outputs = [(-1, "不使用输出")] + [(i, output.get_port_name(i)) for i in range(output.get_port_count())]
        del input, output  # 及时释放资源
        return {'inputs': inputs, 'outputs': outputs}

    def _import_log(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("MIDI Logs", "*.txt *.log"), ("All Files", "*.*")]
        )
        if not file_path:
            return
        
        try:
            with open(file_path, 'r') as f:
                log_content = f.read()
            
            if not log_content:
                messagebox.showerror("转换错误", "日志文件为空")
                self._log_message("日志文件为空", is_error=True)
                return
            
            # 移除人工参数，使用日志自带参数
            try:
                mid = MidiCoreThread.parse_midi_log(log_content)
            except ValueError as e:  # 捕获解析异常
                self._log_message(str(e), is_error=True)
                messagebox.showerror("转换错误", str(e))
                return
            
            save_path = filedialog.asksaveasfilename(
                defaultextension=".mid",
                filetypes=[("MIDI Files", "*.mid"), ("All Files", "*.*")],
                initialdir=self.path_var.get()
            )
            if save_path:
                mid.save(save_path)
                messagebox.showinfo("转换成功", f"文件已保存至：{save_path}")

        except Exception as e:
            messagebox.showerror("转换错误", f"发生错误：{str(e)}")
            self._log_message(f"日志转换失败：{str(e)}", is_error=True)

if __name__ == "__main__":
    root = tk.Tk()
    app = MidiRecorderApp(root)
    root.mainloop()