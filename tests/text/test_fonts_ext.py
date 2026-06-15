"""Unit tests for the Linux/custom font-directory extensions of FontFiles."""

import os

from pptx.text.fonts import FontFiles


class DescribeCustomFontDirectories:
    def it_is_empty_when_env_var_unset(self, monkeypatch):
        monkeypatch.delenv("PYTHON_PPTX_FONT_DIRECTORY", raising=False)
        assert FontFiles._custom_font_directories() == []

    def it_splits_the_env_var_on_the_os_path_separator(self, monkeypatch):
        value = os.pathsep.join(["/opt/fonts", "/srv/brand"])
        monkeypatch.setenv("PYTHON_PPTX_FONT_DIRECTORY", value)
        assert FontFiles._custom_font_directories() == ["/opt/fonts", "/srv/brand"]

    def it_drops_empty_segments(self, monkeypatch):
        value = os.pathsep.join(["/opt/fonts", "", "/srv/brand"])
        monkeypatch.setenv("PYTHON_PPTX_FONT_DIRECTORY", value)
        assert FontFiles._custom_font_directories() == ["/opt/fonts", "/srv/brand"]


class DescribeLinuxFontDirectories:
    def it_includes_the_system_locations(self):
        dirs = FontFiles._linux_font_directories()
        assert "/usr/share/fonts" in dirs
        assert "/usr/local/share/fonts" in dirs

    def it_includes_user_locations_when_HOME_set(self, monkeypatch):
        monkeypatch.setenv("HOME", "/home/fbar")
        dirs = FontFiles._linux_font_directories()
        assert os.path.join("/home/fbar", ".local", "share", "fonts") in dirs
        assert os.path.join("/home/fbar", ".fonts") in dirs


class DescribeFontDirectoriesComposition:
    def it_prepends_custom_dirs_to_platform_defaults(self, monkeypatch):
        monkeypatch.setattr("pptx.text.fonts.sys.platform", "linux")
        monkeypatch.setenv("PYTHON_PPTX_FONT_DIRECTORY", "/opt/fonts")
        dirs = FontFiles._font_directories()
        assert dirs[0] == "/opt/fonts"
        assert "/usr/share/fonts" in dirs

    def it_falls_back_to_custom_dirs_on_unknown_os(self, monkeypatch):
        monkeypatch.setattr("pptx.text.fonts.sys.platform", "freebsd13")
        monkeypatch.setenv("PYTHON_PPTX_FONT_DIRECTORY", "/opt/fonts")
        assert FontFiles._font_directories() == ["/opt/fonts"]
