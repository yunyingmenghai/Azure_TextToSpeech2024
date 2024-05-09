import os
import tkinter as tk
import configparser
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter.scrolledtext import ScrolledText
from datetime import datetime
import azure.cognitiveservices.speech as speechsdk
from azure.cognitiveservices.speech.audio import AudioOutputConfig
from azure.cognitiveservices.speech import SpeechSynthesisOutputFormat
from azure.cognitiveservices.speech import SpeechConfig

def unique_filename(extension):
    date_str = datetime.now().strftime("%Y%m%d")
    timestamp = datetime.now().strftime("%H%M%S")
    filename = f"speech_output_{date_str}_{timestamp}{extension}"
    return filename

class VoiceSynthesisApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("七木AI语音合成器")
        self.minsize(600, 500)

        # 初始化语音选择下拉菜单
        self.voice_var = tk.StringVar()
        self.voice_var.set("zh-CN-XiaoxiaoMultilingualNeural")  # 默认选择
        voices = ["zh-CN-XiaoxiaoMultilingualNeural",
                  "zh-CN-XiaochenMultilingualNeural",
                  "zh-CN-XiaoyuMultilingualNeural",
                  "zh-CN-YunyiMultilingualNeural"]
        
        self.voice_menu = ttk.Combobox(self, textvariable=self.voice_var, values=voices, state="readonly")
        self.voice_menu.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")

        # 初始化文本输入框
        self.text_input = ScrolledText(self, height=15, wrap=tk.WORD, width=70)
        self.text_input.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        # 文本输入框的提示信息
        self.prompt_text = "输入文本（最多3000字）"
        self.text_input.insert(tk.END, self.prompt_text)

        # 初始化字数统计标签
        self.word_count_label = tk.Label(self, text=f"字数: {len(self.prompt_text)}", anchor="e")
        self.word_count_label.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        # 绑定文本输入框的修改事件以更新字数统计
        self.text_input.bind("<KeyRelease>", self.update_word_count)

        # 绑定文本输入框的键盘事件，以便在输入时删除提示文本
        self.text_input.bind("<Key>", self.remove_prompt)

        # 初始化保存路径输入框
        self.save_path_entry = tk.Entry(self, width=40)
        self.save_path_entry.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

        # 创建保存路径选择按钮
        self.save_path_button = tk.Button(self, text="选择保存文件夹", command=self.select_save_path)
        self.save_path_button.grid(row=2, column=2, padx=10, pady=10)

        # 初始化配置Azure订阅密钥和区域的按钮
        self.config_button = tk.Button(self, text="配置Azure订阅密钥和区域", command=self.show_config_dialog)
        self.config_button.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

        # 初始化语音合成按钮，初始状态为禁用
        self.synthesize_button = tk.Button(self, text="转换为语音", state='disabled', command=self.synthesize_speech)
        self.synthesize_button.grid(row=4, columnspan=2, padx=10, pady=10)

        # 初始化SpeechConfig，需要用户配置后才能使用
        self.speech_config = None

        # 存储用户输入的订阅密钥和区域
        self.subscription_key = None
        self.region = None

        # 确保窗口在屏幕的中心
        self.update_idletasks()
        self.geometry(f"{self.winfo_width()}x{self.winfo_height()}+{self.winfo_screenwidth() // 2 - self.winfo_width() // 2}"
                      f"+{self.winfo_screenheight() // 2 - self.winfo_height() // 2}")
        # 检测并加载Azure订阅密钥和区域
        self.load_or_prompt_config()

    def load_or_prompt_config(self):
        config_path = "config.ini"
        # 检测配置文件是否存在
        if not os.path.exists(config_path):
            self.prompt_for_config_and_save()
        else:
            self.load_config(config_path)

    def prompt_for_config_and_save(self):
        # 弹出对话框让用户输入Azure订阅密钥和区域
        input_text = simpledialog.askstring(
            "Azure配置",
            "请输入Azure订阅密钥和区域。",
            parent=self
        )
        if input_text:
            try:
                self.subscription_key, self.region = input_text.strip().split()
                if self.subscription_key and self.region:
                    self.save_config("config.ini", self.subscription_key, self.region)
                    self.speech_config = SpeechConfig(subscription=self.subscription_key, region=self.region)
                    self.synthesize_button.config(state='normal')  # 更新语音合成按钮状态
                else:
                    messagebox.showerror("错误", "输入格式不正确，请输入Azure订阅密钥和区域（用空格分隔）。")
            except ValueError:
                messagebox.showerror("错误", "输入格式不正确，请输入Azure订阅密钥和区域（用空格分隔）。")

    def save_config(self, config_path, subscription_key, region):
        # 保存配置到文件
        config = configparser.ConfigParser()
        config['Azure'] = {'subscription_key': subscription_key, 'region': region}
        with open(config_path, 'w') as configfile:
            config.write(configfile)

    def load_config(self, config_path):
        # 从文件加载配置
        config = configparser.ConfigParser()
        config.read(config_path)
        try:
            self.subscription_key = config.get('Azure', 'subscription_key')
            self.region = config.get('Azure', 'region')
            self.speech_config = SpeechConfig(subscription=self.subscription_key, region=self.region)
            self.synthesize_button.config(state='normal')  # 启用语音合成按钮
        except configparser.NoSectionError as e:
            messagebox.showerror("错误", f"配置文件格式不正确: {e}")

    def update_word_count(self, event=None):
        text = self.text_input.get("1.0", tk.END).strip()
        self.word_count_label.config(text=f"字数: {len(text)}")

    def remove_prompt(self, event):
        if self.text_input.get("1.0", "end-1c") == self.prompt_text:
            self.text_input.delete("1.0", "end-1c")
        self.update_word_count()

    def select_save_path(self):
        folder_path = filedialog.askdirectory(title="选择保存音频的文件夹")
        if folder_path:
            self.save_path_entry.delete(0, tk.END)
            self.save_path_entry.insert(0, folder_path)

    def show_config_dialog(self):
        # 弹出对话框让用户输入Azure订阅密钥和区域
        input_text = simpledialog.askstring(
            "Azure配置",
            f"请输入Azure订阅密钥和区域（用空格分隔）。\n\n留空则使用当前配置。",
            parent=self
        )
        # 如果用户提供了输入，则更新配置
        if input_text:
            try:
                self.subscription_key, self.region = input_text.strip().split()
                if self.subscription_key and self.region:
                    self.speech_config = SpeechConfig(subscription=self.subscription_key, region=self.region)
                    self.synthesize_button.config(state='normal')  # 启用语音合成按钮
            except ValueError:
                messagebox.showerror("错误", "输入格式不正确，请输入Azure订阅密钥和区域（用空格分隔）。")
        else:
            # 用户没有输入，保持当前配置
            pass

    def synthesize_speech(self):
        # 检查是否已配置订阅密钥和区域
        if not self.speech_config:
            messagebox.showerror("错误", "请先配置Azure订阅密钥和区域。")
            return

        text = self.text_input.get("1.0", tk.END).strip()
        if len(text) > 3000:
            messagebox.showerror("错误", "输入的文本超过了3000字的限制。")
            return
        save_path = self.save_path_entry.get()
        if not save_path:
            messagebox.showerror("错误", "请先选择一个保存音频的文件夹。")
            return

        save_path = os.path.join(save_path, unique_filename(".wav"))
        self.synthesize_speech_to_file(text, save_path)

    def synthesize_speech_to_file(self, text, save_path):
        if not self.speech_config:
            return

        # 设置用户选择的语音
        self.speech_config.speech_synthesis_voice_name = self.voice_var.get()

        # 设置自定义音频格式
        self.speech_config.set_speech_synthesis_output_format(SpeechSynthesisOutputFormat.Riff48Khz16BitMonoPcm)

        audio_config = AudioOutputConfig(filename=save_path)

        speech_synthesizer = speechsdk.SpeechSynthesizer(self.speech_config, audio_config)

        result = speech_synthesizer.speak_text_async(text).get()

        if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
            messagebox.showinfo("完成", f"语音合成成功，文件已保存到 {save_path}")
        elif result.reason == speechsdk.ResultReason.Canceled:
            cancellation_details = result.cancellation_details
            messagebox.showerror("取消", f"语音合成取消: {cancellation_details.reason}")
            if cancellation_details.reason == speechsdk.CancellationReason.Error:
                messagebox.showerror("错误", f"错误详情: {cancellation_details.error_details}")

# 确保使用正确的语法来判断是否是主程序
if __name__ == "__main__":
    app = VoiceSynthesisApp()
    app.mainloop()