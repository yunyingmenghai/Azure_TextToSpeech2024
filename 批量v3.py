import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, Text, Scrollbar
import pandas as pd
from datetime import datetime
import azure.cognitiveservices.speech as speechsdk
from azure.cognitiveservices.speech.audio import AudioOutputConfig
from azure.cognitiveservices.speech import SpeechConfig, SpeechSynthesisOutputFormat
from configparser import ConfigParser

class BatchVoiceSynthesisApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("批量语音合成器")
        self.minsize(800, 600)

        # 初始化文本输入框
        self.text_input = tk.Entry(self, width=70)
        self.text_input.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        # 初始化配置Azure订阅密钥和区域的按钮
        self.config_button = tk.Button(self, text="配置Azure订阅密钥和区域", command=self.show_config_dialog)
        self.config_button.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

        # 初始化批量合成按钮，初始状态为禁用
        self.synthesize_button = tk.Button(self, text="批量生成语音", state='disabled', command=self.synthesize_speech)
        self.synthesize_button.grid(row=1, columnspan=1, padx=10, pady=10)


        # 存储用户输入的订阅密钥和区域
        self.subscription_key = None
        self.region = None

        # 确保窗口在屏幕的中心
        self.update_idletasks()
        self.geometry(f"{self.winfo_width()}x{self.winfo_height()}+{self.winfo_screenwidth() // 2 - self.winfo_width() // 2}"
                      f"+{self.winfo_screenheight() // 2 - self.winfo_height() // 2}")

        # 检测并加载Azure订阅密钥和区域
        self.load_or_prompt_config()
                # 初始化日志输出框
        self.log_text = Text(self, wrap=tk.WORD, height=10)
        self.log_text.grid(row=6, column=0, columnspan=3, padx=10, pady=(0, 10), sticky="nsew")

        # 添加滚动条
        self.log_scroll = Scrollbar(self, command=self.log_text.yview)
        self.log_scroll.grid(row=6, column=3, rowspan=1, padx=(0, 10), pady=(0, 10), sticky="ns")

        # 日志文本框与滚动条联动
        self.log_text.config(yscrollcommand=self.log_scroll.set)

        # ... [省略之后的代码，如初始化SpeechConfig等] ...

    def log(self, message):
        # 输出日志信息
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.config(state=tk.DISABLED)

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
        config = ConfigParser()
        config['Azure'] = {'subscription_key': subscription_key, 'region': region}
        with open(config_path, 'w') as configfile:
            config.write(configfile)

    def load_config(self, config_path):
        # 从文件加载配置
        config = ConfigParser()
        config.read(config_path)
        try:
            self.subscription_key = config.get('Azure', 'subscription_key')
            self.region = config.get('Azure', 'region')
            self.speech_config = SpeechConfig(subscription=self.subscription_key, region=self.region)
            self.synthesize_button.config(state='normal')  # 启用语音合成按钮
        except Exception as e:
            messagebox.showerror("错误", f"配置文件格式不正确: {e}")

    def show_config_dialog(self):
        input_text = simpledialog.askstring(
            "Azure配置",
            "请输入Azure订阅密钥和区域（用空格分隔）。\n\n留空则使用当前配置。",
            parent=self
        )
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

        # 选择包含话术的Excel文件
        file_path = filedialog.askopenfilename(
            title="选择包含话术的Excel文件",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not file_path:
            return

        # 读取话术
        scripts = self.load_scripts(file_path)

        # 选择保存音频的文件夹
        save_dir = filedialog.askdirectory(title="选择保存音频的文件夹")
        if not save_dir:
            return

        self.process_scripts(scripts, save_dir)

    def load_scripts(self, file_path):
        df = pd.read_excel(file_path)
        scripts = df[['文件名', '话术']]  # 假设话术在Excel的第一列
        return scripts.values.tolist()

    def process_scripts(self, scripts, save_dir):
        voice_name = "zh-CN-XiaoxiaoMultilingualNeural"  # 选择一个语音类型
        for script in scripts:
            idx, text = script
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            file_name = f"script_{idx}_{timestamp}.wav"
            file_path = os.path.join(save_dir, file_name)

            self.synthesize_speech_to_file(text, file_path, voice_name)

    def synthesize_speech_to_file(self, text, save_path, voice_name):
        if not self.speech_config:
            return

        
        # 设置用户选择的语音
        self.speech_config.speech_synthesis_voice_name = voice_name

        # 设置自定义音频格式
        self.speech_config.set_speech_synthesis_output_format(SpeechSynthesisOutputFormat.Riff48Khz16BitMonoPcm)

        audio_config = AudioOutputConfig(filename=save_path)

        speech_synthesizer = speechsdk.SpeechSynthesizer(self.speech_config, audio_config)

        result = speech_synthesizer.speak_text_async(text).get()

        if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
            self.log(f"语音合成成功，文件已保存到 {save_path}")
        else:
            self.log(f"语音合成失败: {result.reason}")


if __name__ == "__main__":
    app = BatchVoiceSynthesisApp()
    app.mainloop()