from aiogram.dispatcher import Dispatcher
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram import executor, types
from aiogram import Bot
from aiogram.dispatcher import Dispatcher, FSMContext
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.types import InputFile
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from main import gen_analitica


storage = MemoryStorage()
TOKEN = '2146493358:AAH4lkALC3NYXoWbWNYxz5M0HkuT0AuIYVo'

bot = Bot(token=TOKEN)
dp = Dispatcher(bot, storage=storage)

upload_btn = ReplyKeyboardMarkup(resize_keyboard=True)
btn = KeyboardButton('/upload')
upload_btn.add(btn)


class FSMXlsx(StatesGroup):
    file_1 = State()
    file_2 = State()

@dp.message_handler(commands=['start'])
async def start_message(message: types.Message):
    await bot.send_message(message.from_user.id, '👋Привет! Нажми /upload чтоб начать', reply_markup=upload_btn)




@dp.message_handler(commands=['upload'], state=None)
async def upload(message: types.Message, state=FSMContext):
    await FSMXlsx.file_1.set()
    await message.reply('Пожалуйста, загрузите Еженедельный Финансовый отчет Wildberries в формате Excel.')

@dp.message_handler(content_types=['document'], state=FSMXlsx.file_1)
async def import_file_1(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    await message.document.download(destination_file=f"/Users/alexsvoloch/Downloads/TG_DOC/{user_id}/file_1.xlsx")
   
    
    
    await message.reply('Теперь загрузите Аналитика карточек товара (14 отчет)')
    await FSMXlsx.next()

@dp.message_handler(content_types=['document'], state=FSMXlsx.file_2)
async def import_file_2(message: types.Message, state: FSMContext):
    await message.reply('Спасибо, ваш отчет будет готов через минуту')
    user_id = message.from_user.id
    await message.document.download(destination_file=f"/Users/alexsvoloch/Downloads/TG_DOC/{user_id}/file_2.xlsx")
    file_1 = f"/Users/alexsvoloch/Downloads/TG_DOC/{user_id}/file_1.xlsx"
    file_2 = f"/Users/alexsvoloch/Downloads/TG_DOC/{user_id}/file_2.xlsx"
    try:
        gen_analitica(file_1, file_2, user_id)
        await bot.send_document(user_id, open(f"/Users/alexsvoloch/Downloads/TG_DOC/{user_id}/Отчет.xlsx", 'rb'))
    except ValueError:
        await message.answer('Извините, что то пошло не так, возможно вы перепутали отчеты')
    
    await state.finish()

async def on_startup(_):
    print('Bot online')


if __name__ == '__main__':
    executor.start_polling(dp, on_startup=on_startup)   