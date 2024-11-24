import asyncio
import sys
import traceback
from datetime import datetime
from typing import Callable

from aiogram import F, Bot, Dispatcher, BaseMiddleware
from aiogram.filters import CommandStart
from aiogram.types import Message, ReplyKeyboardMarkup, KeyboardButton, ReactionTypeEmoji

from switcher import windows_switcher

BOT_TOKEN = '!!! ТУТ ТОКЕН ВАШЕГО БОТА ИЗ @BotFather !!!'

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()


@dp.message(CommandStart())
async def _start_(message: Message):
    await message.answer(
        '➡️ Презентер (переключение слайдов во время презентации)',
        reply_markup=ReplyKeyboardMarkup(keyboard=[
            [KeyboardButton(text='➡️ Следующий слайд')],
            [KeyboardButton(text='↩️ Предыдущий слайд')]
        ]))


@dp.message(F.text == '➡️ Следующий слайд')
async def _next_(message: Message):
    result = windows_switcher(key='right', target_window='Демонстрация PowerPoint')
    await message.react(reaction=[ReactionTypeEmoji(emoji=['😱', '🤷', '🤷', '👍'][result])])
    await asyncio.sleep(10)
    await message.delete()


@dp.message(F.text == '↩️ Предыдущий слайд')
async def _prev_(message: Message):
    result = windows_switcher(key='left', target_window='Демонстрация PowerPoint')
    await message.react(reaction=[ReactionTypeEmoji(emoji=['😱', '🤷', '🤷', '👍'][result])])
    await asyncio.sleep(10)
    await message.delete()


class MessageMiddleware(BaseMiddleware):
    async def __call__(self, handler: Callable, message: Message, data: dict):
        user = message.from_user
        print(f'[{datetime.now()}] [@{user.username} «{user.full_name}» ({user.id})]: {message.html_text}')
        result = None
        try:
            result = await handler(message, data)
        except Exception as exception:
            print(f'{type(exception).__name__}: {exception}\n{traceback.format_exc(-1)}', file=sys.stderr)
            await message.answer('⚠️ Произошла непредвиденная ошибка')
        return result


dp.message.middleware(MessageMiddleware())


async def run():
    try:
        bot_info = (await bot.get_me())
        print(f'Бот «{bot_info.full_name}» @{bot_info.username} работает')
        await dp.start_polling(bot)
    finally:
        await bot.session.close()


if __name__ == '__main__':
    asyncio.run(run())
