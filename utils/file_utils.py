import logging
import os
import shutil

logger = logging.getLogger(__name__)


def remove(*paths: str) -> None:
    for path in paths:
        try:
            if os.path.exists(path):
                if os.path.isfile(path):
                    os.remove(path)
                elif os.path.isdir(path):
                    shutil.rmtree(path)
        except Exception as e:
            logger.warning(f"Could not remove {path}: {e}")


def clean_temp_files(temp_file: str | None = None) -> None:
    if temp_file:
        remove(temp_file)
    remove(os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Temp', 'gen_py'))
