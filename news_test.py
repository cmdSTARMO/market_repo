import requests as rq
from bs4 import BeautifulSoup
from datetime import datetime
import pandas as pd
import time
import os
import json
import random

# Define the Excel file path
excel_file_path = 'news_output.xlsx'

# 飞书机器人类
class FeishuBot:
    def __init__(self, webhook_url):
        self.webhook_url = webhook_url

    def send_card_message(self, card_content):
        headers = {
            "Content-Type": "application/json; charset=utf-8",
        }
        payload = {
            "msg_type": "interactive",
            "card": card_content
        }
        response = rq.post(url=self.webhook_url, data=json.dumps(payload), headers=headers)
        return response.json()

# 飞书机器人 Webhook URL
webhook_url = 'https://open.feishu.cn/open-apis/bot/v2/hook/22296611-1bc9-40a6-ba02-57b864a62bdd'
bot = FeishuBot(webhook_url)

# Function to check if the Excel file exists and read the existing data
def read_existing_data():
    if os.path.exists(excel_file_path):
        return pd.read_excel(excel_file_path)
    else:
        return pd.DataFrame(columns=['pubtime', 'title', 'summary', 'url', 'publisher'])

# Function to save new data to the Excel file
def save_to_excel(new_articles):
    existing_data = read_existing_data()
    existing_titles = set(existing_data['title'].tolist())  # Convert existing titles to a set for fast lookup

    updated_data = existing_data.copy()

    # Append new articles if they are not duplicates
    for article in new_articles:
        if article['title'] not in existing_titles:
            print(f"Adding new article: {article['title']}")
            updated_data = pd.concat([pd.DataFrame([article]), updated_data], ignore_index=True)
            existing_titles.add(article['title'])  # Add the new title to the set

            # 飞书新闻
            # 构建卡片消息内容
            # 构建卡片消息内容
            card_content = {
                "config": {
                    "wide_screen_mode": True
                },
                "header": {
                    "title": {
                        "tag": "plain_text",
                        "content": f"{article['title']}"
                    },
                    "subtitle": {
                        "tag": "plain_text", # 固定值 plain_text。
                        "content": f"{datetime.now().strftime('%H:%M:%S')}", # 副标题内容。
                    },
                    "text_tag_list": [
                        {
                            "tag": "text_tag",
                            "text": {
                                "tag": "plain_text",
                                "content": "📰要闻"
                            },
                            "color": "indigo"
                        }
                    ],
                    "template": "Yellow"  # 设置顶部为浅yellow色
                },
                "elements": [
                    {
                        "tag": "div",
                        "text": {
                            "tag": "lark_md",
                            "content": f"**摘要：**\n{article['summary']}"
                        }
                    },
                    {
                        "tag": "markdown",
                        "content": f"由<person id = '94dae5e3' show_name = true show_avatar = true style = 'normal'></person>订阅，<link icon='link-copy_outlined' url={article['url']} pc_url='' ios_url='' android_url=''>原文链接</link>。"
                    },
                    {
                        "tag": "hr"
                    },
                    {
                        "tag": "note",
                        "elements": [
                            {
                                "tag": "plain_text",
                                "content": f"数据来源: {article['publisher']}"
                            }
                        ]
                    }
                ]
            }

            # 发送卡片消息
            bot.send_card_message(card_content)

        else:
            print(f"Already have: {article['title']}")

    # Save the updated data back to the Excel file
    updated_data.to_excel(excel_file_path, index=False)

# Main function to run the process in a loop
while True:
    # Get content from Yahoo Finance
    response = rq.get("https://finance.yahoo.com/topic/latest-news/")
    response_text = response.text
    print(1)
    print(response.text)
    # Parse the content using BeautifulSoup
    soup = BeautifulSoup(response_text, 'html.parser')
    fin_stream = soup.find('div', {'id': 'mrt-node-Fin-Stream', 'data-locator': 'subtree-root'})

    articles = []
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if fin_stream:
        for article in fin_stream.find_all('li', class_='js-stream-content'):
            try:
                if 'ad' in article.get('class', []):
                    continue

                title_elem = article.find('h3')
                summary_elem = article.find('p')
                link_elem = article.find('a', class_='mega-item-header-link')
                publisher_elem = article.find('div', {'class': 'C(#959595) Fz(11px) D(ib) Mb(6px)'})

                article_data = {
                    'pubtime': current_time,
                    'title': title_elem.text.strip() if title_elem else None,
                    'summary': summary_elem.text.strip() if summary_elem else None,
                    'url': link_elem['href'] if link_elem else None,
                    'publisher': publisher_elem.text.split('•')[0].strip() if publisher_elem else None,
                }

                if article_data['title'] and article_data['url']:
                    articles.append(article_data)
            except Exception as e:
                print(f"Error processing article: {e}")

    if articles:
        save_to_excel(articles)
    print(articles)
    random_delay = random.uniform(2, 10)
    time.sleep(30 + random_delay)