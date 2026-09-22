#!/usr/bin/env python3
"""Local Telegram bot (long polling).

On Render the bot is part of the web app and receives updates by webhook.
"""
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(BASE, "backend"))

from telegram_bot import run_polling

if __name__ == "__main__":
    run_polling()
