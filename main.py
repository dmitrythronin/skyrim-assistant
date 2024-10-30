import os
import xlsx
import discord
from dotenv import load_dotenv
from discord.ext import commands

intents = discord.Intents.default()
intents.message_content = True
intents.reactions = True
intents.guilds = True
bot = commands.Bot(command_prefix='>',
                   sync_commands_debug=True,
                   sync_commands=True,
                   intents=intents)


@bot.event
async def on_ready():
    print(f'Bot Name: {bot.user}')
    try:
        synced = await bot.tree.sync()
        print(f'Добавлено {len(synced)} команд')
    except Exception as e:
        print(e)


def search_in_table(table, message):
    try:
        return xlsx.search_table(table, message)
    except Exception as e:
        return "К сожалению, при обработке значения ячейки произошла ошибка"


@bot.tree.command(name="reflyem", description="Поиск по базе Reflyem (в разработке)")
async def reflyem(interaction: discord.Interaction, message: str):
    result = search_in_table(xlsx.REFLYEM_TABLE, message)
    await interaction.response.send_message(result)


@bot.tree.command(name="refl", description="Поиск по базе Reflyem (в разработке)")
async def refl(interaction: discord.Interaction, message: str):
    result = search_in_table(xlsx.REFLYEM_TABLE, message)
    await interaction.response.send_message(result)


@bot.tree.command(name="rfad", description="Поиск по базе RfaD (в разработке)")
async def rfad(interaction: discord.Interaction, message: str):
    result = search_in_table(xlsx.RFAD_TABLE, message)
    await interaction.response.send_message(result)


def start_bot():
    load_dotenv()
    token = os.getenv("TOKEN")
    proxy_url = os.getenv("PROXY_URL")
    print("Запуск бота...")
    client = discord.Client(proxy=proxy_url, intents=discord.Intents.all())
    client.run(token)


# Запуск бота через прокси
start_bot()
