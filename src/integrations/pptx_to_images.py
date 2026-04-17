"""PPTX → PNG 이미지 변환 (Windows PowerPoint COM).

분석용으로 슬라이드를 실제 이미지로 보기 위함.
"""
import sys
from pathlib import Path
import win32com.client


def export(pptx_path: Path, out_dir: Path, width=1600):
    out_dir.mkdir(parents=True, exist_ok=True)
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True  # 필수 (COM 제약)
    try:
        prs = app.Presentations.Open(str(pptx_path.resolve()),
                                      ReadOnly=True, WithWindow=False)
        for i, slide in enumerate(prs.Slides, 1):
            out = out_dir / f"slide_{i:03d}.png"
            slide.Export(str(out.resolve()), "PNG", width, int(width * 9 / 16))
            print(f"  [{i:>3}/{len(prs.Slides)}] {out.name}")
        prs.Close()
    finally:
        app.Quit()


if __name__ == "__main__":
    src = Path(sys.argv[1])
    dst = Path(sys.argv[2])
    export(src, dst)
    print(f"완료: {dst}")
