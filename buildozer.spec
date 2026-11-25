[app]
title = Лист Таблицы
package.name = listspreadsheet
package.domain = org.dmitriyfil
source.dir = .
source.include_exts = py,kv,png,jpg,atlas,md,svg,json
icon.filename = %(source.dir)s/assets/icon.png
version = 0.1.0
requirements = python3,kivy,openpyxl,plyer,xlrd
orientation = portrait
fullscreen = 0
android.api = 34
android.minapi = 21
android.ndk = 25b
android.archs = armeabi-v7a
android.accept_sdk_license = True
android.allow_backup = True
android.private_storage = True
p4a.branch = master

[buildozer]
log_level = 2
warn_on_root = 1
