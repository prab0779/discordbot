import discord
from discord.ext import commands
import pandas as pd
import difflib

# Load the data from your file
file_path = r"C:\DiscordBot\values.xlsx.xlsx"  # Ensure the file is in the same directory
data = pd.read_excel(file_path)

# Inspect and print the column names to debug
print("Columns in DataFrame:", data.columns)

# Function to search for item details
def find_item(item_name):
    item_name = item_name.lower()
    match = data[data['Item name'].str.lower().str.contains(item_name, na=False)]  # Correct column name
    if not match.empty:
        row = match.iloc[0]
        return f"**Item name**: {row['Item name']}\n" \
               f"**Demand**: {row['Demand (out of 10)']}/10\n" \
               f"**Value**: {row['Value']}\n" \
               f"**Rate of change**: {row['rate of change']}"
    else:
        return "Item not found. Please check the name and try again."

# Bot setup
intents = discord.Intents.default()
intents.message_content = True  # This allows the bot to read messages
bot = commands.Bot(command_prefix="!", intents=intents)

# Ensure the columns are stripped of spaces
data.columns = data.columns.str.strip()

# Bot events
@bot.event
async def on_ready():
    print(f"Logged in as {bot.user}")

@bot.command(name="value")
async def value(ctx, *, item_name: str):
    response = find_item(item_name)
    await ctx.send(response)

@bot.event
async def on_message(message):
    print(f"Received message: {message.content}")
    await bot.process_commands(message)

# Run the bot
bot.run('INPUT UR TOKEN')  # Replace with your actual bot token
