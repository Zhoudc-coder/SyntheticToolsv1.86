import os
import sys
import json
from pathlib import Path

CONFIG_FILENAME = "config.json"
DEFAULT_RETENTION_DAYS = 12
DEFAULT_AUTO_EXPORT_BOXES = 0
DEFAULT_AUTO_CHECK_PREV_DAYS = 1
DEFAULT_SPECIAL_MODE = 1
DEFAULT_IGNORE_BOX_FORMAT = 1
DEFAULT_IGNORE_SN_FORMAT = 1
DEFAULT_CHECK_DUPLICATES = 0
DEFAULT_SOURCE_DIR = "\\\\192.168.95.106\\a\\M2177-Production Data\\待出库"


def _get_default_log_dir() -> Path:
    env_dir = os.environ.get('MERGE_LOG_DIR')
    if env_dir:
        return Path(env_dir)
    if getattr(sys, 'frozen', False):
        base_dir = Path(sys.executable).parent
    else:
        base_dir = Path(__file__).parent
    return base_dir / 'merge_log'


def _find_config_file() -> Path | None:
    if getattr(sys, 'frozen', False):
        exe_dir = Path(sys.executable).parent
    else:
        exe_dir = Path(__file__).parent
    config_path = exe_dir / CONFIG_FILENAME
    if config_path.exists():
        return config_path
    home_config = Path.home() / CONFIG_FILENAME
    if home_config.exists():
        return home_config
    return None


def _load_config() -> dict:
    config_path = _find_config_file()
    if not config_path:
        return {}
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def _normalize_unc_path(path_str: str) -> str:
    """规范化 Windows UNC 网络路径"""
    path_str = path_str.strip()
    if path_str.startswith('\\\\'):
        return path_str
    if path_str.startswith('\\') and not path_str.startswith('\\\\'):
        return '\\' + path_str
    return path_str


def _load_config_log_dir() -> Path | None:
    config = _load_config()
    log_dir_str = config.get('log_dir')
    if log_dir_str:
        normalized = _normalize_unc_path(log_dir_str)
        return Path(normalized)
    return None


def _load_config_retention_days() -> int:
    config = _load_config()
    retention = config.get('retention_days', DEFAULT_RETENTION_DAYS)
    try:
        days = int(retention)
        if days > 0:
            return days
    except (ValueError, TypeError):
        pass
    return DEFAULT_RETENTION_DAYS


def _load_config_bool_setting(key: str, default: int) -> bool:
    config = _load_config()
    val = config.get(key, default)
    return str(val) == '1'


def get_source_dir() -> str:
    """
    从配置文件读取 source_dir（待出库文件夹的默认路径）。
    返回规范化后的字符串，未设置时返回空字符串。
    """
    config = _load_config()
    val = config.get('source_dir', DEFAULT_SOURCE_DIR)
    if isinstance(val, str):
        return _normalize_unc_path(val)
    return ''


def set_source_dir(path_str: str) -> bool:
    """
    将 source_dir 写入配置文件。返回是否成功。
    """
    path_str = (path_str or '').strip()

    config_path = _find_config_file()
    if config_path is None:
        try:
            _ensure_config_file(_get_default_log_dir())
        except Exception:
            pass
        config_path = _find_config_file()
        if config_path is None:
            return False

    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
    except Exception:
        config_data = {}

    config_data['source_dir'] = path_str

    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4, ensure_ascii=False)
        return True
    except Exception:
        return False


# 配置文件字段的期望顺序（与"设置"面板选项排序一致）
_CONFIG_ORDER = [
    'log_dir',
    'retention_days',
    'auto_export_boxes',
    'auto_check_prev_days',
    'special_mode_default',
    'ignore_box_format',        # 新增：忽略箱码格式
    'ignore_sn_format',         # 新增：忽略SN格式
    'check_duplicates',
    'source_dir'
]

_CONFIG_DEFAULTS = {
    'log_dir': None,            # 运行时由 _ensure_config_file 传入
    'retention_days': DEFAULT_RETENTION_DAYS,
    'auto_export_boxes': DEFAULT_AUTO_EXPORT_BOXES,
    'auto_check_prev_days': DEFAULT_AUTO_CHECK_PREV_DAYS,
    'special_mode_default': DEFAULT_SPECIAL_MODE,
    'ignore_box_format': DEFAULT_IGNORE_BOX_FORMAT,
    'ignore_sn_format': DEFAULT_IGNORE_SN_FORMAT,
    'check_duplicates': DEFAULT_CHECK_DUPLICATES,
    'source_dir': DEFAULT_SOURCE_DIR
}


def _ensure_config_file(default_log_dir: Path):
    """
    如果配置文件不存在，则创建一个包含默认配置的文件；
    若存在但缺少某些字段，则补充默认值并按期望顺序重排。
    """
    defaults = dict(_CONFIG_DEFAULTS)
    defaults['log_dir'] = str(default_log_dir)

    config_path = _find_config_file()
    if config_path is None:
        if getattr(sys, 'frozen', False):
            config_dir = Path(sys.executable).parent
        else:
            config_dir = Path(__file__).parent
        config_path = config_dir / CONFIG_FILENAME
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(defaults, f, indent=4, ensure_ascii=False)
        except Exception:
            home_config = Path.home() / CONFIG_FILENAME
            try:
                with open(home_config, 'w', encoding='utf-8') as f:
                    json.dump(defaults, f, indent=4, ensure_ascii=False)
            except Exception:
                pass
    else:
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)

            changed = False
            new_config = {}
            for k in _CONFIG_ORDER:
                if k in config_data:
                    new_config[k] = config_data[k]
                else:
                    new_config[k] = defaults[k]
                    changed = True
            for k, v in config_data.items():
                if k not in new_config:
                    new_config[k] = v
                    changed = True

            if changed:
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(new_config, f, indent=4, ensure_ascii=False)
        except Exception:
            pass


def _get_log_dir() -> Path:
    env_dir = os.environ.get('MERGE_LOG_DIR')
    if env_dir:
        return Path(env_dir)
    config_dir = _load_config_log_dir()
    if config_dir:
        return config_dir
    return _get_default_log_dir()


LOG_DIR = _get_log_dir()
DB_PATH = LOG_DIR / 'merge_log.db'
RETENTION_DAYS = _load_config_retention_days()
AUTO_EXPORT_BOXES = _load_config_bool_setting('auto_export_boxes', DEFAULT_AUTO_EXPORT_BOXES)
AUTO_CHECK_PREV_DAYS = _load_config_bool_setting('auto_check_prev_days', DEFAULT_AUTO_CHECK_PREV_DAYS)
SPECIAL_MODE_DEFAULT = _load_config_bool_setting('special_mode_default', DEFAULT_SPECIAL_MODE)
IGNORE_BOX_FORMAT_DEFAULT = _load_config_bool_setting('ignore_box_format', DEFAULT_IGNORE_BOX_FORMAT)
IGNORE_SN_FORMAT_DEFAULT = _load_config_bool_setting('ignore_sn_format', DEFAULT_IGNORE_SN_FORMAT)
CHECK_DUPLICATES_DEFAULT = _load_config_bool_setting('check_duplicates', DEFAULT_CHECK_DUPLICATES)
SOURCE_DIR = get_source_dir()
XLSX_SUFFIX = '.xlsx'
CSV_SUFFIX = '.csv'


def ensure_log_dir() -> Path:
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        raise RuntimeError(
            f"无法创建日志目录 {LOG_DIR}，请检查：\n"
            f"1. 路径是否正确（网络共享盘是否已连接？）\n"
            f"2. 是否有写入权限\n"
            f"3. 可修改配置文件 config.json 中的 log_dir 项，或设置环境变量 MERGE_LOG_DIR\n"
            f"原始错误：{e}"
        )
    _ensure_config_file(LOG_DIR)
    return LOG_DIR