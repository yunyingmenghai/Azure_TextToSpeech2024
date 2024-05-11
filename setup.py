from distutils.core import setup
import py2exe

setup(
    name='MyApp',
    version='1.0',
    description='批量转语音',
    windows=['批量合成语音v3.py'],
    options={
        'py2exe': {
            'bundle_files': 1,  # 将所有文件打包到一个 `.exe` 文件中，不生成外部 .egg 文件
            'compressed': True,
            'optimize': 2,  # 优化编译
        }
    }
)