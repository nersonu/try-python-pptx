from pathlib import Path
from typing import Final

import pptx

from pptx.presentation import Presentation


PROJECT_ROOT_PATH: Final[Path] = Path(__file__).parents[1]


def create_presentation() -> Presentation:
    """以下の違いに注意
    - pptx.Presentation(): プレゼンテーションを作る関数
    - pptx.presentation.Presentation: Presetantion オブジェクト
    """
    return pptx.Presentation()


if __name__ == "__main__":
    presentation = create_presentation()
    presentation.save(PROJECT_ROOT_PATH / "dst.pptx")
