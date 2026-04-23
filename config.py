import json
import os
import sys


def _get_config_path():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "config.json")


def _get_default_output_dir():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "output")


def load_config():
    config_path = _get_config_path()
    defaults = {
        "output_mode": "original",
        "custom_output_dir": _get_default_output_dir(),
    }
    if not os.path.exists(config_path):
        return defaults
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            saved = json.load(f)
        defaults.update(saved)
        return defaults
    except Exception:
        return defaults


def save_config(config):
    config_path = _get_config_path()
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        return False


def get_output_dir(source_file=None):
    config = load_config()
    mode = config.get("output_mode", "original")

    if mode == "custom":
        custom_dir = config.get("custom_output_dir", "").strip()
        if custom_dir:
            try:
                os.makedirs(custom_dir, exist_ok=True)
                return custom_dir
            except Exception:
                pass

    if source_file and os.path.isfile(source_file):
        return os.path.dirname(os.path.abspath(source_file))

    return _get_default_output_dir()


def validate_output_dir(path):
    if not path or not path.strip():
        return False, "目录路径不能为空"
    path = path.strip()
    try:
        os.makedirs(path, exist_ok=True)
        test_file = os.path.join(path, ".markitdown_test")
        with open(test_file, "w") as f:
            f.write("test")
        os.remove(test_file)
        return True, ""
    except PermissionError:
        return False, f"没有权限访问目录: {path}"
    except OSError as e:
        return False, f"目录路径无效: {e}"
