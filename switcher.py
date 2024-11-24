import enum
import sys
import time
import traceback

import keyboard
import win32com.client
import win32con
import win32gui


class SwitchStatus(enum.IntEnum):
    error = 0
    not_found = 1
    forcibly_pressed = 2
    successful = 3


def windows_switcher(key: str, target_window: str = None, *, is_forced: bool = True) -> SwitchStatus:
    """
    Переключает фокус окон Windows на указанное и имитирует нажатие на клавишу.

    :param key: Клавиша, на которую нужно нажать.
    :param target_window: Часть или весь заголовок целевого окна.
    :param is_forced: Если окно не найдено, то всё равно имитирует нажатие клавиши.
    :return: Статус операции
    """
    try:
        # Окно не указано
        if not isinstance(target_window, str):
            if is_forced:
                keyboard.send(key)  # Принудительно имитируем нажатие клавиши
                return SwitchStatus.forcibly_pressed
            return SwitchStatus.not_found

        # Поиск активного окна
        active_window = win32gui.GetWindowText(win32gui.GetForegroundWindow())

        # Если активное окно является целевым
        if target_window in active_window:
            keyboard.send(key)
            return SwitchStatus.successful

        # Ищем целевое окно
        _list_, win_list = list(), list()

        def function(hwnd, _):
            _list_.append(win32gui.GetWindowText(hwnd))
            if target_window in win32gui.GetWindowText(hwnd):
                win_list.append(hwnd)

        win32gui.EnumWindows(function, None)

        # Если список найденных окон пуст
        if len(win_list) == 0 and is_forced:
            if is_forced:
                keyboard.send(key)
                return SwitchStatus.forcibly_pressed
            return SwitchStatus.not_found

        # Запоминаем старое окно
        old_window = win32gui.GetForegroundWindow()

        # Переключаемся на целевое окно
        """
        Отправляем левый альтернативный ключ. По какой-то причине windows действительно хочет,
        чтобы была нажата клавиша alt перед тем, как выдвинуть на передний план любое другое окно
        
        Информация о ошибке:
        https://stackoverflow.com/questions/14295337/win32gui-setactivewindow-error-the-specified-procedure-could-not-be-found
        """
        shell = win32com.client.Dispatch('WScript.Shell')
        shell.SendKeys('%')

        win32gui.ShowWindow(win_list[0], win32con.SW_SHOW)
        win32gui.SetForegroundWindow(win_list[0])

        # Имитируем нажатие клавиши
        time.sleep(0.1)
        keyboard.send(key)

        # Возвращаемся к прежнему окну
        win32gui.ShowWindow(old_window, win32con.SW_SHOW)
        win32gui.SetForegroundWindow(old_window)

        return SwitchStatus.successful

    except Exception as exception:
        print(f'{type(exception).__name__}: {exception}\n{traceback.format_exc(-1)}', file=sys.stderr)
        return SwitchStatus.error
