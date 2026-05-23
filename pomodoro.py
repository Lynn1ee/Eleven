"""
一个基于 Python tkinter 的桌面番茄钟应用。

番茄工作法的核心规则：
  - 25 分钟专注工作（一个番茄钟）
  - 5 分钟短休息
  - 每完成 4 个番茄钟后，进行一次 15 分钟长休息
  - 循环往复
"""

import tkinter as tk
import math
import time
import threading

# Windows 平台使用 winsound 发出提示音，其他平台回退到系统响铃
try:
    import winsound
    HAS_WINSOUND = True
except ImportError:
    HAS_WINSOUND = False


class PomodoroTimer:
    """
    番茄钟主类，封装所有 UI 和计时逻辑。

    状态机说明：
      mode 可取三个值: "work" / "short_break" / "long_break"
      running: 是否正在计时
      paused:  是否已暂停（running=True 但计时暂停）
    """

    # ==================== 时长配置（分钟） ====================
    WORK_MIN = 25            # 专注时长
    SHORT_BREAK_MIN = 5      # 短休息时长
    LONG_BREAK_MIN = 15      # 长休息时长
    SESSIONS_BEFORE_LONG = 4 # 每完成几个番茄钟后触发长休息

    # ==================== 配色方案 ====================
    COLOR_WORK = "#e74c3c"        # 专注模式主色（番茄红）
    COLOR_WORK_BG = "#fdf2f2"     # 专注模式背景色（浅红）
    COLOR_BREAK = "#27ae60"       # 休息模式主色（薄荷绿）
    COLOR_BREAK_BG = "#f0faf4"    # 休息模式背景色（浅绿）
    COLOR_BG = "#ffffff"          # 窗口背景色
    COLOR_TEXT = "#2c3e50"        # 主文字颜色（深灰蓝）
    COLOR_SUBTEXT = "#7f8c8d"     # 次要文字颜色（灰）
    COLOR_BTN_BG = "#f8f9fa"      # 按钮默认背景
    COLOR_BTN_HOVER = "#e9ecef"   # 按钮悬停背景
    COLOR_PROGRESS_BG = "#ecf0f1" # 进度条背景
    COLOR_DOT_DONE = "#2c3e50"    # 已完成圆点颜色
    COLOR_DOT_TODO = "#dcdde1"    # 未完成圆点颜色

    # ================================================================
    #  初始化：创建窗口、设置初始状态、构建 UI
    # ================================================================
    def __init__(self):
        # ----- 创建主窗口 -----
        self.root = tk.Tk()
        self.root.title("Pomodoro Timer")
        self.root.geometry("400x540")            # 固定窗口大小 400×540
        self.root.resizable(False, False)        # 禁止拖拽改变窗口大小

        # 先让窗口管理器计算实际尺寸，避免高 DPI 下布局偏移
        self.root.update_idletasks()
        self.root.configure(bg=self.COLOR_BG)

        # 清除默认 icon（Windows 下可能显示为空白图标）
        self.root.iconbitmap(default="")

        # 用代码绘制简易番茄图标（不需要外部 ico 文件）
        self._set_window_icon()

        # ----- 初始化计时状态 -----
        self.remaining_sec = self.WORK_MIN * 60  # 剩余秒数，初始为 25 分钟
        self.total_sec = self.WORK_MIN * 60      # 总秒数，用于计算进度比例
        self.running = False                     # 是否正在运行
        self.paused = False                      # 是否已暂停
        self.mode = "work"                       # 当前模式: work / short_break / long_break
        self.session_count = 0                   # 已完成的番茄钟数量
        self._after_id = None                    # tkinter after() 返回的定时器 ID，用于取消定时
        self._sound_stop_event = None            # 用于停止提示音线程的 Event

        # ----- 构建界面并刷新显示 -----
        self._build_ui()
        self._update_display()

        # 点击窗口关闭按钮时，安全退出
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ================================================================
    #  窗口图标：纯代码绘制一个 32×32 的番茄图标
    # ================================================================
    def _set_window_icon(self):
        """用像素计算绘制简易番茄图标（红色果实 + 绿色蒂）。"""
        try:
            icon = tk.PhotoImage(width=32, height=32)
            for y in range(32):
                for x in range(32):
                    # 以 (16, 20) 为圆心，计算当前像素到圆心的距离
                    cx, cy = x - 16, y - 20
                    dist = math.sqrt(cx * cx + cy * cy)
                    if dist < 10:
                        # 圆形范围内用红色（番茄果实）
                        icon.put("#e74c3c", (x, y))
                    elif 6 < x < 26 and y > 22 and y < 28:
                        # 底部中间用绿色（番茄蒂）
                        icon.put("#27ae60", (x, y))
                    else:
                        # 其余区域透明
                        icon.put("#ffffff00", (x, y))
            self.root.iconphoto(True, icon)
        except Exception:
            # 图标绘制失败不影响主功能，静默忽略
            pass

    # ================================================================
    #  UI 构建：将所有控件拼装到窗口上
    # ================================================================
    def _build_ui(self):
        """
        窗口布局（从上到下）:
          top_frame     → 模式标签 + 圆形画布（进度环）+ 倒计时文字
          session_frame → 会话计数 + 进度圆点
          btn_frame     → 开始 / 暂停 / 重置 按钮
          bottom_frame  → 跳过按钮 + 置顶复选框
        """

        # ----- 顶部区域：模式标签、进度环画布（含倒计时）-----
        self.top_frame = tk.Frame(self.root, bg=self.COLOR_BG)
        self.top_frame.pack(pady=(20, 0))

        self.mode_label = tk.Label(
            self.top_frame, text="FOCUS", font=("Segoe UI", 13, "bold"),
            bg=self.COLOR_BG, fg=self.COLOR_WORK
        )
        self.mode_label.pack()

        # Canvas 画布：用 create_text() 在圆心绘制倒计时，彻底消除绝对定位错位
        self.canvas_size = 280
        self.canvas = tk.Canvas(
            self.top_frame, width=self.canvas_size, height=self.canvas_size,
            bg=self.COLOR_BG, highlightthickness=0
        )
        self.canvas.pack(pady=(8, 0))

        # 先创建倒计时文字（后续绘制需要用 _timer_text_id 做 tag_raise）
        self._timer_text_id = self.canvas.create_text(
            self.canvas_size / 2, self.canvas_size / 2,
            text="25:00", font=("Segoe UI", 44, "bold"),
            fill=self.COLOR_TEXT, anchor="center", tags="timer"
        )

        # 再绘制完整圆环
        self._draw_progress_ring(1.0)

        # ----- 中部区域：会话计数 + 进度圆点 -----
        self.session_frame = tk.Frame(self.root, bg=self.COLOR_BG)
        self.session_frame.pack(pady=(16, 0))

        self.session_label = tk.Label(
            self.session_frame, text="Sessions: 0",
            font=("Segoe UI", 11), bg=self.COLOR_BG, fg=self.COLOR_SUBTEXT
        )
        self.session_label.pack()

        # 4 个圆点，表示当前周期内完成了几个番茄钟
        self.dots_frame = tk.Frame(self.session_frame, bg=self.COLOR_BG)
        self.dots_frame.pack(pady=(6, 0))

        self.dot_labels = []
        for i in range(self.SESSIONS_BEFORE_LONG):
            dot = tk.Label(
                self.dots_frame, text="●", font=("Segoe UI", 14),
                bg=self.COLOR_BG, fg=self.COLOR_DOT_TODO  # 初始为未完成状态（灰色）
            )
            dot.pack(side="left", padx=4)
            self.dot_labels.append(dot)

        # ----- 按钮区域：开始、暂停、重置 -----
        self.btn_frame = tk.Frame(self.root, bg=self.COLOR_BG)
        self.btn_frame.pack(pady=(20, 0))

        # 开始按钮（初始为红色，与专注模式一致）
        self.start_btn = tk.Button(
            self.btn_frame, text="▶  Start", font=("Segoe UI", 12, "bold"),
            bg=self.COLOR_WORK, fg="white", activebackground="#c0392b",
            activeforeground="white", relief="flat", padx=28, pady=10,
            cursor="hand2", command=self._start
        )
        self.start_btn.pack(side="left", padx=6)

        # 暂停按钮（初始禁用，计时开始后才可点击）
        self.pause_btn = tk.Button(
            self.btn_frame, text="⏸  Pause", font=("Segoe UI", 12, "bold"),
            bg=self.COLOR_BTN_BG, fg=self.COLOR_TEXT,
            activebackground=self.COLOR_BTN_HOVER, relief="flat",
            padx=28, pady=10, cursor="hand2", command=self._pause,
            state="disabled"
        )
        self.pause_btn.pack(side="left", padx=6)

        # 重置按钮
        self.reset_btn = tk.Button(
            self.btn_frame, text="↺  Reset", font=("Segoe UI", 12, "bold"),
            bg=self.COLOR_BTN_BG, fg=self.COLOR_TEXT,
            activebackground=self.COLOR_BTN_HOVER, relief="flat",
            padx=28, pady=10, cursor="hand2", command=self._reset
        )
        self.reset_btn.pack(side="left", padx=6)

        # ----- 底部区域：跳过按钮 + 窗口置顶开关 -----
        self.bottom_frame = tk.Frame(self.root, bg=self.COLOR_BG)
        self.bottom_frame.pack(pady=(16, 10))

        self.skip_btn = tk.Button(
            self.bottom_frame, text="Skip →", font=("Segoe UI", 10),
            bg=self.COLOR_BG, fg=self.COLOR_SUBTEXT, relief="flat",
            cursor="hand2", command=self._skip, bd=0,
            activebackground=self.COLOR_BG
        )
        self.skip_btn.pack(side="left", padx=6)

        # "始终置顶" 复选框，默认勾选
        self.always_on_top_var = tk.BooleanVar(value=True)
        self.always_on_top_cb = tk.Checkbutton(
            self.bottom_frame, text="Always on top",
            variable=self.always_on_top_var, font=("Segoe UI", 10),
            bg=self.COLOR_BG, fg=self.COLOR_SUBTEXT,
            selectcolor=self.COLOR_BG, activebackground=self.COLOR_BG,
            cursor="hand2", command=self._toggle_always_on_top
        )
        self.always_on_top_cb.pack(side="left", padx=12)
        self.root.attributes("-topmost", True)  # 初始即置顶

        # 为三个主按钮绑定鼠标悬停变色效果
        self._set_hover(self.start_btn, self.COLOR_WORK, "#c0392b")
        self._set_hover(self.pause_btn, self.COLOR_BTN_BG, self.COLOR_BTN_HOVER)
        self._set_hover(self.reset_btn, self.COLOR_BTN_BG, self.COLOR_BTN_HOVER)

    # ================================================================
    #  辅助方法：按钮悬停变色
    # ================================================================
    def _set_hover(self, btn, normal, hover):
        """绑定按钮的鼠标进入/离开事件，颜色存储在 btn 属性上，支持后续动态修改。"""
        btn._hover_normal = normal
        btn._hover_hover = hover

        def on_enter(e):
            if e.widget["state"] != "disabled":
                e.widget.configure(bg=e.widget._hover_hover)

        def on_leave(e):
            if e.widget["state"] != "disabled":
                e.widget.configure(bg=e.widget._hover_normal)

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)

    # ================================================================
    #  进度环绘制：在 Canvas 上画出多层环形进度条
    # ================================================================
    def _draw_progress_ring(self, ratio):
        """
        绘制环形进度条。

        参数:
          ratio: 0.0 ~ 1.0，表示剩余时间比例。
                 ratio=1.0 → 完整圆环
                 ratio=0.0 → 圆环消失

        实现思路：
          - 画 3 层同心圆弧，每层半径略有不同，形成层次感
          - 使用 create_arc() 从正上方（90°）顺时针绘制
          - 内层画一个实心圆作为背景，让数字清晰可见
          - 绘制完毕后将计时文字抬升到最上层
        """
        self.canvas.delete("ring")  # 清空旧图形（tag="ring" 的所有元素）

        cx = self.canvas_size / 2   # 画布中心 X
        cy = self.canvas_size / 2   # 画布中心 Y
        r = 108                     # 基准半径
        width = 6                   # 弧线粗细

        color = self.COLOR_WORK if self.mode == "work" else self.COLOR_BREAK

        # 画 3 层同心圆弧，形成"光晕"效果
        for i in range(3):
            rr = r + width + 8 + i * 4  # 每层半径递增 4px

            deg = 360 * ratio if ratio > 0 else 0

            if ratio > 0.999:
                self.canvas.create_oval(
                    cx - rr, cy - rr, cx + rr, cy + rr,
                    outline=color, width=2, tags="ring"
                )
            else:
                self.canvas.create_arc(
                    cx - rr, cy - rr, cx + rr, cy + rr,
                    start=90, extent=-deg,
                    outline=color, width=2, style="arc", tags="ring"
                )

        # 内层实心圆：作为倒计时数字的背景
        inner_r = r - 4
        bg_color = self.COLOR_WORK_BG if self.mode == "work" else self.COLOR_BREAK_BG
        self.canvas.create_oval(
            cx - inner_r, cy - inner_r, cx + inner_r, cy + inner_r,
            fill=bg_color, outline="", tags="ring"
        )

        # 确保倒计时文字渲染在最上层，不被进度环遮挡
        self.canvas.tag_raise(self._timer_text_id)

    # ================================================================
    #  刷新显示：更新倒计时数字、进度环、标题栏、按钮状态
    # ================================================================
    def _update_display(self):
        """每次计时变化时调用，同步所有 UI 元素。"""
        # 格式化时间为 MM:SS
        mins = self.remaining_sec // 60
        secs = self.remaining_sec % 60
        time_str = f"{mins:02d}:{secs:02d}"

        # 更新 Canvas 上的倒计时文字（不再用 place 定位的 Label）
        self.canvas.itemconfigure(self._timer_text_id, text=time_str)

        # 更新环形进度条
        ratio = self.remaining_sec / self.total_sec if self.total_sec > 0 else 0
        self._draw_progress_ring(ratio)

        # 更新标题栏，方便在任务栏看到剩余时间
        title_mode = "Focus" if self.mode == "work" else "Break"
        self.root.title(f"{time_str} - {title_mode}")

        self._update_button_states()

    # ================================================================
    #  按钮状态管理：根据运行状态启用/禁用相应按钮
    # ================================================================
    def _update_button_states(self):
        """
        按钮状态逻辑：
          - 未运行:            [▶ Start] 可用,  [⏸ Pause] 禁用
          - 运行中（未暂停）:   [▶ Start] 禁用,  [⏸ Pause] 可用
          - 已暂停:            [▶ Resume] 可用, [⏸ Pause] 禁用
        """
        work_color = self.COLOR_WORK if self.mode == "work" else self.COLOR_BREAK

        if self.running and not self.paused:
            self.start_btn.configure(state="disabled")
            self.pause_btn.configure(state="normal")
        elif self.running and self.paused:
            self.start_btn.configure(
                state="normal", text="▶  Resume", bg=work_color
            )
            self.start_btn._hover_normal = work_color
            self.start_btn._hover_hover = "#c0392b" if self.mode == "work" else "#1e8449"
            self.pause_btn.configure(state="disabled")
        else:
            self.start_btn.configure(
                state="normal", text="▶  Start", bg=work_color
            )
            self.start_btn._hover_normal = work_color
            self.start_btn._hover_hover = "#c0392b" if self.mode == "work" else "#1e8449"
            self.pause_btn.configure(state="disabled")

    # ================================================================
    #  模式切换时更新 UI 配色
    # ================================================================
    def _set_mode_ui(self):
        """根据当前模式（work / break）切换标签文字、按钮颜色和悬停颜色。"""
        if self.mode == "work":
            self.mode_label.configure(text="FOCUS", fg=self.COLOR_WORK)
            self.start_btn.configure(bg=self.COLOR_WORK)
            self.start_btn._hover_normal = self.COLOR_WORK
            self.start_btn._hover_hover = "#c0392b"
        else:
            self.mode_label.configure(text="BREAK", fg=self.COLOR_BREAK)
            self.start_btn.configure(bg=self.COLOR_BREAK)
            self.start_btn._hover_normal = self.COLOR_BREAK
            self.start_btn._hover_hover = "#1e8449"

    # ================================================================
    #  核心计时逻辑：每秒触发一次
    # ================================================================
    def _tick(self):
        """
        计时器的心跳函数，每秒被 tkinter 的 after() 回调一次。

        流程：
          1. 检查是否仍在运行且未暂停
          2. 剩余秒数 -1，刷新显示
          3. 若倒计时到 0 → 触发 _timer_finished()
          4. 否则 → 注册下一秒的 after() 回调
        """
        if not self.running or self.paused:
            return

        if self.remaining_sec > 0:
            self.remaining_sec -= 1
            self._update_display()

        if self.remaining_sec <= 0:
            self._timer_finished()   # 时间到！
        else:
            # 1000ms 后再次调用自身，形成每秒循环
            self._after_id = self.root.after(1000, self._tick)

    # ================================================================
    #  计时结束后的处理：播放声音、切换模式、更新圆点
    # ================================================================
    def _timer_finished(self, skipped=False):
        """
        当前时段倒计时归零时触发。

        参数:
          skipped: True 表示用户跳过（跳过时不增加 session_count）

        模式切换规则：
          专注结束 → 短休息（每 4 次中第 4 次为长休息）
          休息结束 → 专注
        """
        self.running = False
        self._after_id = None       # 清除可能已过期的定时器 ID，防止后续误操作

        self._play_sound()          # 播放提示音（异步，不阻塞 UI）

        if self.mode == "work":
            # 完成一个番茄钟（跳过不计入）
            if not skipped:
                self.session_count += 1
                self.session_label.configure(text=f"Sessions: {self.session_count}")
                self._update_dots()

            # 判断是长休息还是短休息
            if self.session_count % self.SESSIONS_BEFORE_LONG == 0:
                self.mode = "long_break"
                self.remaining_sec = self.LONG_BREAK_MIN * 60
                self.total_sec = self.LONG_BREAK_MIN * 60
            else:
                self.mode = "short_break"
                self.remaining_sec = self.SHORT_BREAK_MIN * 60
                self.total_sec = self.SHORT_BREAK_MIN * 60
        else:
            # 休息结束，回到专注模式
            self.mode = "work"
            self.remaining_sec = self.WORK_MIN * 60
            self.total_sec = self.WORK_MIN * 60

        self._set_mode_ui()
        self._update_display()
        self._flash_window()        # 让窗口闪烁引起注意

    # ================================================================
    #  进度圆点更新
    # ================================================================
    def _update_dots(self):
        """
        更新 4 个圆点的颜色，表示当前周期内的完成进度。

        例如：完成了 3 个番茄钟 → 前 3 个圆点变深色，第 4 个保持浅色
        当完成第 4 个时 → 4 个全亮，然后下一轮重置
        """
        completed = self.session_count % self.SESSIONS_BEFORE_LONG
        # 如果正好是 4 的倍数（且 > 0），则显示 4 个全满
        if completed == 0 and self.session_count > 0:
            completed = self.SESSIONS_BEFORE_LONG

        for i in range(self.SESSIONS_BEFORE_LONG):
            if i < completed:
                self.dot_labels[i].configure(fg=self.COLOR_DOT_DONE)
            else:
                self.dot_labels[i].configure(fg=self.COLOR_DOT_TODO)

    # ================================================================
    #  声音提示：在单独线程中播放，避免阻塞 UI
    # ================================================================
    def _play_sound(self):
        """
        播放三段提示音。

        Windows：在 daemon 线程中用 winsound.Beep 播放，不阻塞 UI。
                  通过 _sound_stop_event 支持提前终止。
        其他平台：用 after() 在主线程依次响铃，避免 Tk 线程安全问题。
        """
        if HAS_WINSOUND:
            self._sound_stop_event = threading.Event()

            def _play():
                for _ in range(3):
                    if self._sound_stop_event.is_set():
                        break
                    try:
                        winsound.Beep(880, 150)
                    except RuntimeError:
                        break  # 无音频设备时静默退出
                    if self._sound_stop_event.is_set():
                        break
                    time.sleep(0.1)
                    try:
                        winsound.Beep(1100, 300)
                    except RuntimeError:
                        break
                    if self._sound_stop_event.is_set():
                        break
                    time.sleep(0.1)

            t = threading.Thread(target=_play, daemon=True)
            t.start()
        else:
            # 在主线程用 after() 调度响铃，不触碰 Tk 线程安全红线
            self._bell_sequence(0)

    def _bell_sequence(self, count):
        """在主线程上用 after() 递归调度响铃，避免线程安全问题。"""
        if count >= 3:
            return
        self.root.bell()
        self.root.after(500, lambda: self._bell_sequence(count + 1))

    # ================================================================
    #  窗口闪烁：计时结束时将窗口置于最前
    # ================================================================
    def _flash_window(self):
        """短暂置顶窗口并获取焦点，2 秒后恢复用户的置顶偏好。"""
        try:
            was_topmost = self.root.attributes("-topmost")
            self.root.attributes("-topmost", True)
            self.root.focus_force()
            # 如果用户原本关闭了置顶，2 秒后恢复
            if not was_topmost:
                self.root.after(2000, lambda: self.root.attributes(
                    "-topmost", self.always_on_top_var.get()))
        except Exception:
            pass

    # ================================================================
    #  按钮事件处理
    # ================================================================
    def _start(self):
        """
        开始 / 继续按钮。

        两种场景：
          1. 从未开始过（running=False）→ 启动新的计时
          2. 已暂停（paused=True）     → 从暂停处继续
        """
        if not self.running:
            # 首次启动
            self.running = True
            self.paused = False
            self._update_button_states()
            self._after_id = self.root.after(1000, self._tick)
        elif self.paused:
            # 从暂停恢复
            self.paused = False
            self._update_button_states()
            self._after_id = self.root.after(1000, self._tick)

    def _pause(self):
        """暂停计时：取消定时回调，保留当前剩余时间。"""
        if self.running and not self.paused:
            self.paused = True
            if self._after_id:
                self.root.after_cancel(self._after_id)  # 取消已注册的定时器
                self._after_id = None
            self._update_button_states()

    def _reset(self):
        """重置当前时段：停止计时和提示音，恢复到该模式的初始时长。"""
        self.running = False
        self.paused = False

        # 停止正在播放的提示音
        if self._sound_stop_event:
            self._sound_stop_event.set()

        if self._after_id:
            try:
                self.root.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

        # 根据当前模式恢复到对应的初始时长
        if self.mode == "work":
            self.remaining_sec = self.WORK_MIN * 60
            self.total_sec = self.WORK_MIN * 60
        elif self.mode == "long_break":
            self.remaining_sec = self.LONG_BREAK_MIN * 60
            self.total_sec = self.LONG_BREAK_MIN * 60
        else:
            self.remaining_sec = self.SHORT_BREAK_MIN * 60
            self.total_sec = self.SHORT_BREAK_MIN * 60

        self._update_display()

    def _skip(self):
        """
        跳过当前时段：直接将剩余时间归零，触发计时结束逻辑。

        跳过的工作时段不计入 session_count，避免误触发长休息。
        """
        if self._after_id:
            try:
                self.root.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        self.running = False
        self.paused = False
        self.remaining_sec = 0
        self._timer_finished(skipped=True)

    def _toggle_always_on_top(self):
        """切换窗口置顶状态。"""
        self.root.attributes("-topmost", self.always_on_top_var.get())

    def _on_close(self):
        """关闭窗口时的清理：停止计时器和提示音，销毁窗口。"""
        self.running = False
        if self._sound_stop_event:
            self._sound_stop_event.set()
        if self._after_id:
            try:
                self.root.after_cancel(self._after_id)
            except Exception:
                pass
        self.root.destroy()

    # ================================================================
    #  启动入口
    # ================================================================
    def run(self):
        """进入 tkinter 主事件循环，显示窗口。"""
        self.root.mainloop()


# ================================================================
#  程序入口
# ================================================================
if __name__ == "__main__":
    app = PomodoroTimer()
    app.run()
