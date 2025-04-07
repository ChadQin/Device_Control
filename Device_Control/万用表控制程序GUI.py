import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from Keysight_34461A import DMM34461A
import pyvisa
import time,threading,logging

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')


class MultimeterGUI:
    def __init__(self, master):
        self.master = master
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)  # 捕获窗口关闭事件
        self.master.iconbitmap('./resources/steam.ico')
        self.multimeter = None  # 延迟初始化
        self.connection_status = False
        self.scan_thread = None
        self.stop_event = threading.Event()  # 使用Event代替布尔标志
        self.available_devices = []
        self.create_widgets()
        self.refresh_devices()  # 启动时自动扫描设备

    def create_widgets(self):
        """创建界面组件（优化布局版本）"""
        # 配置主窗口布局权重
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_rowconfigure(3, weight=1)  # 输出区域所在行可扩展

        # 设备选择区域
        self.device_frame = ttk.LabelFrame(self.master, text="设备选择")
        self.device_frame.grid(row=0, column=0, padx=10, pady=5, sticky="ew")
        self.device_frame.grid_columnconfigure(0, weight=1)  # 下拉框可扩展

        self.device_combobox = ttk.Combobox(self.device_frame, state="readonly")
        self.device_combobox.grid(row=0, column=0, padx=5, pady=2, sticky="ew")

        self.btn_refresh = ttk.Button(
            self.device_frame,
            text="刷新设备",
            command=self.refresh_devices
        )
        self.btn_refresh.grid(row=0, column=1, padx=5, pady=2)

        # 状态显示区域
        self.status_frame = ttk.LabelFrame(self.master, text="状态信息")
        self.status_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")  # 修正行号为1

        self.status_label = ttk.Label(self.status_frame, text="未连接设备", foreground="red")
        self.status_label.pack(side=tk.LEFT, padx=5, fill='x', expand=True)

        # 控制按钮区域
        self.control_frame = ttk.LabelFrame(self.master, text="设备控制")
        self.control_frame.grid(row=2, column=0, padx=10, pady=5, sticky="ew")

        self.btn_connect = ttk.Button(self.control_frame, text="连接设备", command=self.connect_device)
        self.btn_connect.grid(row=0, column=0, padx=5, pady=2)

        self.btn_disconnect = ttk.Button(
            self.control_frame,
            text="断开连接",
            command=self.disconnect_device,
            state=tk.DISABLED
        )
        self.btn_disconnect.grid(row=0, column=1, padx=5, pady=2)

        # 设备信息按钮
        self.btn_info = ttk.Button(
            self.control_frame,
            text="设备信息",
            command=self.show_device_info,
            state=tk.DISABLED
        )
        self.btn_info.grid(row=0, column=2, padx=5, pady=2)

        # 测量功能区域（防拉伸优化版）
        self.measure_frame = ttk.LabelFrame(self.master, text="测量功能")
        self.measure_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")

        # 配置行列权重（关键修改）
        self.measure_frame.columnconfigure([0, 1], weight=1, uniform="measure_col")  # 均匀列宽
        self.measure_frame.rowconfigure([0, 1, 2], weight=0)  # 关闭行权重，保持自然高度

        # 定义按钮样式（固定最小尺寸）
        style = ttk.Style()
        style.configure("Fixed.TButton",
                        width=12,  # 固定字符宽度
                        anchor="center",  # 文字居中
                        padding=5)  # 内边距

        # 统一使用固定样式
        btn_style = "Fixed.TButton"
        btn_padx = 3
        btn_pady = 3

        # 第一行：电压测量
        self.btn_dc_voltage = ttk.Button(
            self.measure_frame,
            text="直流电压",
            command=self.measure_dc_voltage,
            state=tk.DISABLED,
            style=btn_style
        )
        self.btn_dc_voltage.grid(row=0, column=0, padx=btn_padx, pady=btn_pady, sticky="ew")

        self.btn_ac_voltage = ttk.Button(
            self.measure_frame,
            text="交流电压",
            command=self.measure_ac_voltage,
            state=tk.DISABLED,
            style=btn_style
        )
        self.btn_ac_voltage.grid(row=0, column=1, padx=btn_padx, pady=btn_pady, sticky="ew")

        # 第二行：电流测量
        self.btn_dc_current = ttk.Button(
            self.measure_frame,
            text="直流电流",
            command=self.measure_dc_current,
            state=tk.DISABLED,
            style=btn_style
        )
        self.btn_dc_current.grid(row=1, column=0, padx=btn_padx, pady=btn_pady, sticky="ew")

        self.btn_ac_current = ttk.Button(
            self.measure_frame,
            text="交流电流",
            command=self.measure_ac_current,
            state=tk.DISABLED,
            style=btn_style
        )
        self.btn_ac_current.grid(row=1, column=1, padx=btn_padx, pady=btn_pady, sticky="ew")

        # 第三行：电阻测量
        self.btn_resistance = ttk.Button(
            self.measure_frame,
            text="电阻",
            command=self.measure_resistance,
            state=tk.DISABLED,
            style=btn_style
        )
        self.btn_resistance.grid(row=2, column=0, columnspan=2, padx=btn_padx, pady=btn_pady, sticky="ew")

        # 主窗口布局约束
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(3, weight=0)  # 固定测量区域高度


        self.scan_frame = ttk.LabelFrame(self.master, text="数据扫描")
        self.scan_frame.grid(row=4, column=0, padx=10, pady=5, sticky="ew")

        # 配置扫描帧的列权重（新增第3列）
        self.scan_frame.columnconfigure(0, weight=1)  # 按钮组向左对齐
        self.scan_frame.columnconfigure(1, weight=0)  # 参数选择固定宽度
        self.scan_frame.columnconfigure(2, weight=0)  # 间隔输入框（新增）
        self.scan_frame.columnconfigure(3, weight=0)  # 状态标签固定宽度

        # 扫描控制按钮组（左侧，保持不变）
        button_frame = ttk.Frame(self.scan_frame)
        button_frame.grid(row=0, column=0, sticky="w")

        self.btn_start_scan = ttk.Button(
            button_frame,
            text="开始扫描",
            command=self.start_scan,
            state=tk.DISABLED
        )
        self.btn_start_scan.pack(side=tk.LEFT, padx=5, pady=2)

        self.btn_stop_scan = ttk.Button(
            button_frame,
            text="停止扫描",
            command=self.stop_scan,
            state=tk.DISABLED
        )
        self.btn_stop_scan.pack(side=tk.LEFT, padx=5, pady=2)

        # 参数选择下拉菜单（中间）
        self.param_frame = ttk.Frame(self.scan_frame)
        self.param_frame.grid(row=0, column=1, sticky="e", padx=(20, 5))

        ttk.Label(self.param_frame, text="测量模式:").pack(side=tk.LEFT)

        self.param_combobox = ttk.Combobox(
            self.param_frame,
            values=["直流电压", "交流电压", "直流电流", "交流电流", "电阻"],
            state="readonly",
            width=12
        )
        self.param_combobox.current(0)  # 默认选择第一个
        self.param_combobox.pack(side=tk.LEFT, padx=5)
        self.selected_param = self.param_combobox.get()  # 实时获取当前选择

        # ===== 新增间隔输入框 =====
        self.interval_frame = ttk.Frame(self.scan_frame)
        self.interval_frame.grid(row=0, column=2, sticky="e", padx=(10, 5))

        ttk.Label(self.interval_frame, text="间隔(秒):").pack(side=tk.LEFT)

        self.interval_value = tk.StringVar(value="1.0")  # 默认1秒
        vcmd = (self.master.register(self.validate_number), '%P')
        self.interval_entry = ttk.Entry(
            self.interval_frame,
            textvariable=self.interval_value,
            width=6,
            validate="key",
            validatecommand=vcmd
        )
        self.interval_entry.pack(side=tk.LEFT, padx=(5, 0))

        # 扫描状态指示器（最右侧，列号改为3）
        self.scan_status = ttk.Label(
            self.scan_frame,
            text="就绪",
            foreground="gray"
        )
        self.scan_status.grid(row=0, column=3, sticky="e", padx=10)

        # 调整主窗口行权重
        self.master.rowconfigure(4, weight=0)

        # 输出信息区域
        self.output_frame = ttk.LabelFrame(self.master, text="输出信息")
        self.output_frame.grid(row=5, column=0, padx=10, pady=5, sticky="nsew")  # 使用nsew填充
        self.output_frame.grid_columnconfigure(0, weight=1)
        self.output_frame.grid_rowconfigure(0, weight=1)

        self.output_text = scrolledtext.ScrolledText(
            self.output_frame,
            wrap=tk.WORD,
            height=10  # 设置固定行数
        )
        self.output_text.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.output_text.configure(font=('Courier New', 9))  # 可选等宽字体

        # 新增退出按钮
        self.btn_exit = ttk.Button(
            self.control_frame,
            text="退出",
            command=self.confirm_exit,
            style="TButton"
        )
        self.btn_exit.grid(row=0, column=3, padx=5, pady=2)  # 新增第4列

        pass



    def update_status(self, message, color="black"):
        """仅更新状态标签"""
        self.status_label.config(text=message, foreground=color)
        pass

    def update_output(self, message):
        """更新输出信息框"""
        self.output_text.insert(tk.END, f"{message}\n")
        self.output_text.see(tk.END)

    def refresh_devices(self):
        """刷新可用设备列表"""
        try:
            rm = pyvisa.ResourceManager()
            self.available_devices = rm.list_resources()
            self.device_combobox["values"] = self.available_devices
            if self.available_devices:
                self.device_combobox.current(0)
                self.update_status(f"发现{len(self.available_devices)}个设备", "blue")
            else:
                self.update_status("未检测到可用设备", "red")
        except Exception as e:
            messagebox.showerror("VISA错误", f"资源管理器初始化失败: {str(e)}")

    def connect_device(self):
        """连接选定设备（添加设备信息按钮控制）"""
        if not self.available_devices:
            messagebox.showwarning("警告", "请先刷新可用设备列表")
            return

        selected_device = self.device_combobox.get()
        if not selected_device:
            return

        try:
            # ============= 保持原有连接逻辑 =============
            self.multimeter = DMM34461A()
            self.multimeter.K34461A = self.multimeter.rm.open_resource(selected_device)
            self.connection_status = True
            self.update_status(f"已连接至 {selected_device}", "green")  # 保持原有状态更新
            self.btn_connect.config(state=tk.DISABLED)
            self.btn_disconnect.config(state=tk.NORMAL)
            self.enable_measure_buttons(True)

            # ============= 新增设备信息按钮控制 =============
            if hasattr(self, 'btn_info'):  # 安全检测防止属性不存在
                self.btn_info.config(state=tk.NORMAL)  # 关键修复点：启用按钮
                self.btn_start_scan.config(state=tk.NORMAL)
                self.scan_status.config(text="就绪", foreground="blue")
            else:
                print("Warning: 设备信息按钮未初始化")


        except Exception as e:
            # ============= 保持原有错误处理 =============
            self.update_status(f"连接失败: {str(e)}", "red")
            messagebox.showerror("连接错误", f"设备连接失败: {str(e)}")

            # ============= 新增错误状态处理 =============
            if hasattr(self, 'btn_info'):
                self.btn_info.config(state=tk.DISABLED)  # 确保失败时保持禁用
                self.btn_start_scan.config(state=tk.DISABLED)

        pass

    def disconnect_device(self):
        try:
            if self.multimeter and self.multimeter.K34461A:
                self.multimeter.K34461A.close()
                self.connection_status = False
                self.update_status("连接已断开", "orange")
                self.btn_connect.config(state=tk.NORMAL)
                self.btn_disconnect.config(state=tk.DISABLED)
                self.enable_measure_buttons(False)
                self.btn_info.config(state=tk.DISABLED)
                self.stop_event.set()
            logging.debug("Device disconnected")
        except Exception as e:
            self.update_status(f"断开失败: {str(e)}", "red")
            self.btn_connect.config(state=tk.NORMAL)
            self.btn_disconnect.config(state=tk.DISABLED)
            logging.error(f"Disconnect error: {str(e)}")

    def show_device_info(self):
        """显示设备详细信息"""
        """显示设备详细信息窗口"""
        if not self.connection_status:
            messagebox.showwarning("警告", "请先连接设备")
            return

        try:
            info_window = tk.Toplevel(self.master)
            info_window.title("设备详细信息")
            info_window.resizable(False, False)

            # 主信息框架
            main_frame = ttk.Frame(info_window)
            main_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

            # 设备基本信息
            basic_info_frame = ttk.LabelFrame(main_frame, text="基本配置")
            basic_info_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

            device = self.multimeter.K34461A
            idn_info = device.query('*IDN?').strip().split(',')

            info_data = [
                ("VISA地址：", device.resource_name),
                ("厂商：", idn_info[0]),
                ("型号：", idn_info[1]),
                ("序列号：", idn_info[2]),
                ("固件版本：", idn_info[3]),
                ("超时设置：", f"{device.timeout} ms"),
                ("接口类型：", device.resource_class)
            ]

            # 动态生成信息标签
            for row, (label, value) in enumerate(info_data):
                ttk.Label(basic_info_frame, text=label, font=('TkDefaultFont', 9, 'bold')) \
                    .grid(row=row, column=0, padx=2, pady=2, sticky='e')
                ttk.Label(basic_info_frame, text=value, font=('TkFixedFont')) \
                    .grid(row=row, column=1, padx=2, pady=2, sticky='w')


            # 底部按钮
            btn_frame = ttk.Frame(main_frame)
            btn_frame.grid(row=2, column=0, pady=10)

            ttk.Button(
                btn_frame,
                text="保存配置",
                command=lambda: self.save_device_config(device)
            ).grid(row=0, column=0, padx=5)

            ttk.Button(
                btn_frame,
                text="关闭",
                command=info_window.destroy
            ).grid(row=0, column=1, padx=5)

        except Exception as e:
            messagebox.showerror("错误", f"获取设备信息失败: {str(e)}")

    def save_device_config(self, device):
        """保存设备配置示例方法"""
        # 这里可以添加实际保存逻辑
        messagebox.showinfo("提示", "配置保存功能待实现")

    def enable_measure_buttons(self, enable):
        """控制测量按钮状态"""
        state = tk.NORMAL if enable else tk.DISABLED
        self.btn_dc_voltage.config(state=state)
        self.btn_ac_voltage.config(state=state)
        self.btn_dc_current.config(state=state)
        self.btn_ac_current.config(state=state)
        self.btn_resistance.config(state=state)

    def measure_dc_voltage(self):
        """测量电压"""
        if not self.connection_status:
            messagebox.showwarning("警告", "请先连接设备")
            return
        try:
            voltage = self.multimeter.get_volt_dc()
            self.update_output(f"直流电压测量值: {voltage:.6f} V")  # 使用update_output
            self.multimeter.local()
        except Exception as e:
            self.update_output(f"测量失败: {str(e)}")
            messagebox.showerror("测量错误", f"电压测量时出错: {str(e)}")

    def measure_ac_voltage(self):
        if not self.connection_status:
            messagebox.showwarning("警告", "请先连接设备")
            return
        try:
            voltage = self.multimeter.get_volt_ac()
            self.update_output(f"交流电压测量值: {voltage:.6f} V")
            self.multimeter.local()
        except Exception as e:
            self.update_output(f"测量失败: {str(e)}")
            messagebox.showerror("测量错误", f"电压测量时出错: {str(e)}")

    def measure_dc_current(self):
        """测量电流"""
        if not self.connection_status:
            messagebox.showwarning("警告", "请先连接设备")
            return

        try:
            current = self.multimeter.get_curr_dc()  # 调用电流测量方法
            self.update_output(f"直流电流测量值: {current:.6f} A")
            self.multimeter.local()
        except Exception as e:
            self.update_output(f"测量失败: {str(e)}")
            messagebox.showerror("测量错误", f"电流测量时出错: {str(e)}")

    def measure_ac_current(self):
        """测量电流"""
        if not self.connection_status:
            messagebox.showwarning("警告", "请先连接设备")
            return

        try:
            current = self.multimeter.get_curr_dc()  # 调用电流测量方法
            self.update_output(f"交流电流测量值: {current:.6f} A")
            self.multimeter.local()
        except Exception as e:
            self.update_output(f"测量失败: {str(e)}")
            messagebox.showerror("测量错误", f"电流测量时出错: {str(e)}")

    def measure_resistance(self):
        """测量电阻"""
        if not self.connection_status:
            messagebox.showwarning("警告", "请先连接设备")
            return

        try:
            resistance = self.multimeter.get_immp()  # 调用电阻测量方法
            print(type(resistance))
            if resistance > 100000000:
                self.update_output("overload Ω")
                self.multimeter.local()
            else:
                self.update_output(f"电阻测量值: {resistance:.6f} Ω")
                self.multimeter.local()
        except Exception as e:
            self.update_output(f"测量失败: {str(e)}")
            messagebox.showerror("测量错误", f"电阻测量时出错: {str(e)}")

    def start_scan(self):
        """启动扫描线程"""
        if not self.connection_status:
            messagebox.showwarning("警告", "请先连接设备")
            return
        if self.scan_thread is None or not self.scan_thread.is_alive():
            self.stop_event.clear()
            self.scan_thread = threading.Thread(target=self.start_scanning)
            self.scan_thread.start()
            self.btn_start_scan.config(state=tk.DISABLED)
            self.btn_stop_scan.config(state=tk.NORMAL)
            self.scan_status.config(text="扫描中...", foreground="green")
        logging.debug("Scan started")

    def stop_scan(self):
        """停止扫描线程"""
        if self.scan_thread and self.scan_thread.is_alive():
            self.stop_event.set()
            self.master.after(100, self.join_scan_thread)
        else:
            self.update_ui_after_stop()
        logging.debug("Stop scan requested")

    def join_scan_thread(self):
        if self.scan_thread and self.scan_thread.is_alive():
            logging.debug("Waiting for scan thread to join...")
            self.scan_thread.join(timeout=2)  # 设置2秒超时
            if self.scan_thread.is_alive():
                logging.warning("Scan thread forced termination")
                # 强制终止线程（谨慎使用）
                self.scan_thread._stop()
            logging.debug("Scan thread joined")
        self.update_ui_after_stop()

    def update_ui_after_stop(self):
        """更新UI状态"""
        self.btn_start_scan.config(state=tk.NORMAL)
        self.btn_stop_scan.config(state=tk.DISABLED)
        self.scan_status.config(text="已停止", foreground="red")
        logging.debug("UI updated after stop")

    def start_scanning(self):
        try:
            while not self.stop_event.is_set():
                meas_param = self.param_combobox.get()
                try:
                    interval = float(self.interval_value.get())
                except ValueError:
                    interval = 1.0

                try:
                    if self.stop_event.is_set():
                        break

                    logging.debug(f"Starting measurement for {meas_param}")
                    start_time = time.time()
                    if meas_param == '直流电压':
                        param = self.multimeter.get_volt_dc()
                    elif meas_param == '交流电压':
                        param = self.multimeter.get_volt_ac()
                    elif meas_param == '直流电流':
                        param = self.multimeter.get_curr_dc()
                    elif meas_param == '交流电流':
                        param = self.multimeter.get_curr_ac()
                    elif meas_param == '电阻':
                        param = self.multimeter.get_immp()
                    else:
                        param = None

                    end_time = time.time()
                    logging.debug(f"{meas_param} took {end_time - start_time:.3f}s")

                    if self.stop_event.is_set():
                        break

                    msg = self.format_measurement(meas_param, param)
                    self.master.after(0, self.update_output, msg)  # 使用update_output

                except Exception as e:
                    self.master.after(0, self.update_output, f"扫描错误: {str(e)}")
                    logging.error(f"Scan error: {str(e)}")
                    break

                time.sleep(interval)
        except Exception as e:
            logging.error(f"Unexpected error in scan thread: {str(e)}")
        finally:
            self.master.after(0, self.update_ui_after_stop)

    def confirm_exit(self):
        """显示确认退出对话框"""
        if messagebox.askyesno("退出确认", "确认要退出程序吗？"):
            self.master.destroy()

    def on_closing(self):
        """处理窗口关闭事件"""
        self.confirm_exit()

    def format_measurement(self, meas_param, param):
        if meas_param == '电阻' and param > 1e8:
            return "电阻测量值: overload Ω"
        elif param is not None:
            return f"{meas_param}测量值: {param:.6f} {'V' if '电压' in meas_param else 'A' if '电流' in meas_param else 'Ω'}"
        else:
            return "未知测量结果"


    def stop_scanning(self):
        """绑定万用表类的停止扫描方法"""
        self.stop_flag = True
        if self.scan_thread is not None and self.scan_thread.is_alive():
            self.scan_thread.join()  # 等待线程结束

        # 更新按钮状态
        self.btn_start_scan.config(state=tk.NORMAL)
        self.btn_stop_scan.config(state=tk.DISABLED)
        self.scan_status.config(text="已停止", foreground="red")

    def validate_number(self, new_value):
        """验证输入是否为有效数字"""
        if new_value == "":
            return True  # 允许清空
        try:
            float(new_value)
            return True
        except ValueError:
            return False

    # def get_interval(self):
    #     """获取间隔时间（秒）"""
    #     try:
    #         self.interval_value = self.interval_value.get()
    #         return float(self.interval_value.get())
    #     except ValueError:
    #         messagebox.showerror("错误", "请输入有效的数字")
    #         return 1.0  # 默认返回1秒


if __name__ == "__main__":
    root = tk.Tk()
    root.title("万用表控制程序")
    root.geometry("800x600")
    app = MultimeterGUI(root)
    root.mainloop()