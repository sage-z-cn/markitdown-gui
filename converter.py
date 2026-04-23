import os
from PyQt5.QtCore import QThread, pyqtSignal


class ConvertThread(QThread):
    progress_signal = pyqtSignal(int)
    status_signal = pyqtSignal(str)
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(str, str, bool, str)

    def __init__(self, file_path, output_dir):
        super().__init__()
        self.file_path = file_path
        self.output_dir = output_dir

    def run(self):
        file_name = os.path.basename(self.file_path)
        self.status_signal.emit("正在转换...")
        self.progress_signal.emit(10)
        self.log_signal.emit(f"开始转换文件: {file_name}")

        try:
            try:
                os.makedirs(self.output_dir, exist_ok=True)
            except PermissionError:
                raise RuntimeError(f"没有权限创建输出目录: {self.output_dir}")
            except OSError as e:
                raise RuntimeError(f"输出目录无效: {e}")

            if not os.path.isdir(self.output_dir):
                raise RuntimeError(f"输出目录不存在或无法访问: {self.output_dir}")

            self.log_signal.emit(f"正在加载文件: {file_name}")
            self.progress_signal.emit(30)

            from markitdown import MarkItDown
            md = MarkItDown()
            result = md.convert(self.file_path)

            self.progress_signal.emit(60)
            self.log_signal.emit(f"文件转换完成，正在保存...")

            base_name = os.path.splitext(file_name)[0]
            output_file = os.path.join(self.output_dir, base_name + ".md")

            counter = 1
            while os.path.exists(output_file):
                output_file = os.path.join(
                    self.output_dir, f"{base_name}_{counter}.md"
                )
                counter += 1

            try:
                with open(output_file, "w", encoding="utf-8") as f:
                    f.write(result.text_content)
            except PermissionError:
                raise RuntimeError(f"没有权限写入文件: {output_file}")
            except OSError as e:
                raise RuntimeError(f"写入文件失败: {e}")

            self.progress_signal.emit(100)
            self.status_signal.emit("转换成功")
            self.log_signal.emit(f"文件已保存至: {output_file}")
            self.finished_signal.emit(self.file_path, output_file, True, "")

        except Exception as e:
            self.progress_signal.emit(100)
            self.status_signal.emit("转换失败")
            error_msg = str(e)
            self.log_signal.emit(f"转换失败: {error_msg}")
            output_file = os.path.join(
                self.output_dir,
                os.path.splitext(file_name)[0] + ".md",
            )
            self.finished_signal.emit(self.file_path, output_file, False, error_msg)
