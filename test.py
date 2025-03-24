from notify_util import FeishuBot

bot = FeishuBot("https://open.feishu.cn/open-apis/bot/v2/hook/22296611-1bc9-40a6-ba02-57b864a62bdd")

try:
    # 模拟出错
    1 / 0
except Exception as e:
    bot.send_card_message(
        content=f"程序运行报错：\n```{str(e)}```",
        title="🚨 报错警告",
        tag_text="错误",
        tag_color="red"
    )
