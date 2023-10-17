import argostranslate.package
import argostranslate.translate

class ArgosTranslateInfo:
    import pathlib

    def __init__(self, from_code: str = 'en', to_code: str = 'jp', download_path: pathlib.Path = None):
        self.from_code: str = from_code
        self.to_code: str = to_code
        self.download_path = download_path

    def do_download(self, install_package: argostranslate.package.AvailablePackage) -> None:
        # 今回使用するパッケージをダウンロードし、そのパスを保存しておく
        self.set_download_path(install_package.download())

    def set_download_path(self, download_path: pathlib.Path) -> None:
        self.download_path = download_path

def do_download(pkg_info: ArgosTranslateInfo) -> None:
    # 使用可能な翻訳パッケージ情報を取得
    available_pkgs: list = argostranslate.package.get_available_packages()

    # 今回使う翻訳パッケージ情報を取得
    install_package: argostranslate.package.AvailablePackage = next(
        filter(
            lambda x: x.from_code == pkg_info.from_code and x.to_code == pkg_info.to_code,
            available_pkgs
        )
    )

    # ダウンロード実施 (パスも保存される)
    pkg_info.do_download(install_package)

def do_translate(text: str, from_code: str = 'en', to_code: str = 'ja') -> str:
    translatedText: str = argostranslate.translate.translate(
        text, from_code, to_code)
    if to_code == 'ja':
        print(translatedText)
    return translatedText


# 以下が基本的な実行手順
# 翻訳したいパターンをlistで指定
pkg_infos: list = [
    ArgosTranslateInfo(from_code='en', to_code='ja'),
]

# Packageインデクス情報をダウンロード確認 (アップデート確認)
argostranslate.package.update_package_index()

# 今回使用するパッケージ情報を取得し、パッケージをダウンロードし、そのパスを保存しておく
for pkg_info in pkg_infos:
    do_download(pkg_info)

# 2回目以降の実行では、pkg_infosのdownload_pathさえあれば、これ以降の手順のみで実行可能

# 今回使用するパッケージを読み込み
for pkg_info in pkg_infos:
    argostranslate.package.install_from_path(pkg_info.download_path)

# Translate
do_translate('polling.')

