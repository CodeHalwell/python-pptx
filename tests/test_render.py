"""Unit-test suite for `pptx.render`.

LibreOffice is mocked out: these tests cover argument plumbing, error
handling, and selection of slide indexes — not the actual rendering.
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import pytest

from pptx import Presentation
from pptx.render import (
    DEFAULT_TIMEOUT_SECONDS,
    ThumbnailRendererError,
    ThumbnailRendererUnavailable,
    render_slide_thumbnail,
    render_slide_thumbnails,
)


@pytest.fixture
def two_slide_prs():
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[5])
    prs.slides.add_slide(prs.slide_layouts[5])
    return prs


def _fake_soffice_run(work_dir: Path, *, num_slides=2, exit_code=0, stderr=b""):
    """Drop fake `*.png` files in the work_dir and return a stub CompletedProcess."""
    from subprocess import CompletedProcess

    def _runner(soffice_bin, deck_path, out_dir, timeout):
        if exit_code == 0:
            for i in range(num_slides):
                (Path(out_dir) / f"slide{i}.png").write_bytes(b"\x89PNG\r\n\x1a\n%d" % i)
        return CompletedProcess(args=[], returncode=exit_code, stdout=b"", stderr=stderr)

    return _runner


class DescribeRenderSlideThumbnails:
    def it_returns_paths_for_every_slide_by_default(self, tmp_path, two_slide_prs):
        with patch("pptx.render.shutil.which", return_value="/usr/bin/soffice"), patch(
            "pptx.render._run_soffice", side_effect=_fake_soffice_run(tmp_path)
        ):
            paths = render_slide_thumbnails(two_slide_prs, out_dir=tmp_path)

        assert len(paths) == 2
        assert all(p.suffix == ".png" for p in paths)

    def it_returns_only_requested_indexes(self, tmp_path, two_slide_prs):
        with patch("pptx.render.shutil.which", return_value="/usr/bin/soffice"), patch(
            "pptx.render._run_soffice", side_effect=_fake_soffice_run(tmp_path)
        ):
            paths = render_slide_thumbnails(
                two_slide_prs, out_dir=tmp_path, slide_indexes=[1]
            )

        assert len(paths) == 1
        assert paths[0].name.endswith("1.png")

    def it_raises_when_index_is_out_of_range(self, tmp_path, two_slide_prs):
        with patch("pptx.render.shutil.which", return_value="/usr/bin/soffice"), patch(
            "pptx.render._run_soffice", side_effect=_fake_soffice_run(tmp_path)
        ):
            with pytest.raises(IndexError, match="out of range"):
                render_slide_thumbnails(
                    two_slide_prs, out_dir=tmp_path, slide_indexes=[5]
                )

    def it_returns_bytes_when_asked(self, tmp_path, two_slide_prs):
        with patch("pptx.render.shutil.which", return_value="/usr/bin/soffice"), patch(
            "pptx.render._run_soffice", side_effect=_fake_soffice_run(tmp_path)
        ):
            data = render_slide_thumbnails(
                two_slide_prs, out_dir=tmp_path, return_bytes=True
            )

        assert all(isinstance(d, bytes) for d in data)
        assert data[0].startswith(b"\x89PNG")

    def it_raises_unavailable_when_soffice_is_missing(self, two_slide_prs):
        with patch("pptx.render.shutil.which", return_value=None):
            with pytest.raises(ThumbnailRendererUnavailable, match="install LibreOffice"):
                render_slide_thumbnails(two_slide_prs)

    def it_raises_when_soffice_exits_non_zero(self, tmp_path, two_slide_prs):
        with patch("pptx.render.shutil.which", return_value="/usr/bin/soffice"), patch(
            "pptx.render._run_soffice",
            side_effect=_fake_soffice_run(tmp_path, exit_code=1, stderr=b"boom"),
        ):
            with pytest.raises(ThumbnailRendererError, match="status 1"):
                render_slide_thumbnails(two_slide_prs, out_dir=tmp_path)

    def it_raises_when_soffice_produces_no_pngs(self, tmp_path, two_slide_prs):
        with patch("pptx.render.shutil.which", return_value="/usr/bin/soffice"), patch(
            "pptx.render._run_soffice", side_effect=_fake_soffice_run(tmp_path, num_slides=0)
        ):
            with pytest.raises(ThumbnailRendererError, match="no PNG output"):
                render_slide_thumbnails(two_slide_prs, out_dir=tmp_path)

    def it_passes_the_default_timeout_through(self, tmp_path, two_slide_prs):
        captured = {}

        def _capture(soffice_bin, deck_path, out_dir, timeout):
            captured["timeout"] = timeout
            return _fake_soffice_run(tmp_path)(soffice_bin, deck_path, out_dir, timeout)

        with patch("pptx.render.shutil.which", return_value="/usr/bin/soffice"), patch(
            "pptx.render._run_soffice", side_effect=_capture
        ):
            render_slide_thumbnails(two_slide_prs, out_dir=tmp_path)

        assert captured["timeout"] == DEFAULT_TIMEOUT_SECONDS

    def it_honours_custom_soffice_bin(self, tmp_path, two_slide_prs):
        captured = {}

        def _capture(soffice_bin, deck_path, out_dir, timeout):
            captured["bin"] = soffice_bin
            return _fake_soffice_run(tmp_path)(soffice_bin, deck_path, out_dir, timeout)

        with patch("pptx.render.shutil.which", return_value="/opt/libre/soffice"), patch(
            "pptx.render._run_soffice", side_effect=_capture
        ):
            render_slide_thumbnails(
                two_slide_prs, out_dir=tmp_path, soffice_bin="/opt/libre/soffice"
            )

        assert captured["bin"] == "/opt/libre/soffice"


class DescribeRenderSlideThumbnail:
    def it_renders_the_specific_slide(self, tmp_path, two_slide_prs):
        slide = two_slide_prs.slides[1]

        with patch("pptx.render.shutil.which", return_value="/usr/bin/soffice"), patch(
            "pptx.render._run_soffice", side_effect=_fake_soffice_run(tmp_path)
        ):
            target = tmp_path / "out" / "slide.png"
            path = render_slide_thumbnail(slide, out_path=target)

        assert path == target
        assert path.exists()

    def it_returns_bytes_when_asked(self, tmp_path, two_slide_prs):
        slide = two_slide_prs.slides[0]

        with patch("pptx.render.shutil.which", return_value="/usr/bin/soffice"), patch(
            "pptx.render._run_soffice", side_effect=_fake_soffice_run(tmp_path)
        ):
            data = render_slide_thumbnail(slide, return_bytes=True)

        assert isinstance(data, bytes)
        assert data.startswith(b"\x89PNG")


class DescribePresentationConvenienceMethods:
    def it_exposes_render_thumbnails_on_presentation(self, tmp_path, two_slide_prs):
        with patch("pptx.render.shutil.which", return_value="/usr/bin/soffice"), patch(
            "pptx.render._run_soffice", side_effect=_fake_soffice_run(tmp_path)
        ):
            paths = two_slide_prs.render_thumbnails(out_dir=tmp_path)

        assert len(paths) == 2

    def it_exposes_render_thumbnail_on_slide(self, tmp_path, two_slide_prs):
        slide = two_slide_prs.slides[0]

        with patch("pptx.render.shutil.which", return_value="/usr/bin/soffice"), patch(
            "pptx.render._run_soffice", side_effect=_fake_soffice_run(tmp_path)
        ):
            path = slide.render_thumbnail(out_path=tmp_path / "s.png")

        assert path.exists()
