[app]
title = Product Quote Sheet
package.name = quoteapp
package.domain = org.quote
source.dir = .
source.include_exts = py,png,jpg,kv,atlas
version = 1.0

requirements = python3,kivy==2.3.0,kivymd==1.2.0,openpyxl,pillow,pyjnius,android

orientation = portrait
fullscreen = 0

android.permissions = CAMERA, READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE
android.api = 33
android.minapi = 21

[buildozer]
log_level = 2
warn_on_root = 1
