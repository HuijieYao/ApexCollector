from PyQt5.QtCore import QSettings

settings = QSettings("config.ini", QSettings.IniFormat)
dir_path = settings.value("SETUP/DIR_PATH")
apart = settings.value("SETUP/APART")
floor_index = eval(settings.value("SETUP/FLOOR_INDEX"))
auto_pack = settings.value("SETUP/AUTO_PACK")
is_changed = False
has_printed = False


def set_dir_path(value):
    global dir_path, is_changed
    if dir_path != value:
        is_changed = True
        dir_path = value
        settings.setValue("SETUP/DIR_PATH", dir_path)


def set_apart(value):
    global apart, is_changed
    if apart != value:
        is_changed = True
        apart = value
        settings.setValue("SETUP/APART", apart)


def set_floor_index(value):
    global floor_index, is_changed
    if floor_index != value:
        is_changed = True
        floor_index = value
        settings.setValue("SETUP/FLOOR_INDEX", floor_index)


def set_auto_pack(value):
    global auto_pack, is_changed
    if auto_pack != value:
        is_changed = True
        auto_pack = value
        settings.setValue("SETUP/AUTO_PACK", auto_pack)


def set_is_changed():
    global is_changed
    is_changed = False if is_changed else True
