"""统一日志：stderr + logs/app.log，避免多实例/终端缓冲导致看不到输出。"""

import logging
import os
import sys

LOG_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "logs")
LOG_FILE = os.path.join(LOG_DIR, "app.log")


def setup_logging(level: int = logging.INFO) -> str:
    """配置根日志，返回日志文件路径。"""
    os.makedirs(LOG_DIR, exist_ok=True)

    root = logging.getLogger()
    if getattr(root, "_auto_v2_configured", False):
        return LOG_FILE

    fmt = logging.Formatter(
        "%(asctime)s %(levelname)s [%(name)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    stderr_handler = logging.StreamHandler(sys.stderr)
    stderr_handler.setFormatter(fmt)

    file_handler = logging.FileHandler(LOG_FILE, mode="a", encoding="utf-8")
    file_handler.setFormatter(fmt)

    root.setLevel(level)
    root.handlers.clear()
    root.addHandler(stderr_handler)
    root.addHandler(file_handler)

    # 让 werkzeug 访问日志走根 logger（避免自带 handler 与根配置不一致）
    wz = logging.getLogger("werkzeug")
    wz.handlers.clear()
    wz.propagate = True
    wz.setLevel(logging.INFO)

    for name in ("urllib3", "httpx", "httpcore", "openai"):
        logging.getLogger(name).setLevel(logging.WARNING)

    root._auto_v2_configured = True  # type: ignore[attr-defined]
    return LOG_FILE
