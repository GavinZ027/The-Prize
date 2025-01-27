import random
import pandas as pd
from openpyxl import load_workbook


def load_participants(file_path="name.xlsx"):
    """从Excel文件加载参与者名单"""
    try:
        df = pd.read_excel(file_path)
        if "姓名" not in df.columns:
            print("错误：Excel文件中缺少'姓名'列")
            return []
        return df["姓名"].tolist()
    except FileNotFoundError:
        print(f"错误：找不到文件 {file_path}")
        return []
    except Exception as e:
        print(f"读取Excel文件时发生错误：{str(e)}")
        return []


def multi_level_draw(participants):
    """执行分级抽奖"""
    # 定义奖项配置
    prizes = {
        "一等奖": {"人数": 1, "颜色": "\033[1;31m"},  # 红色
        "二等奖": {"人数": 2, "颜色": "\033[1;33m"},  # 黄色
        "三等奖": {"人数": 3, "颜色": "\033[1;32m"}  # 绿色
    }

    # 检查总人数
    total_winners = sum(p["人数"] for p in prizes.values())
    if len(participants) < total_winners:
        print(f"错误：参与者不足（需至少{total_winners}人，当前{len(participants)}人）")
        return None

    # 执行抽奖
    result = {}
    remaining = participants.copy()

    for prize, config in prizes.items():
        # 抽取获奖者
        winners = random.sample(remaining, config["人数"])
        result[prize] = winners
        # 移除已中奖者
        remaining = [p for p in remaining if p not in winners]

    return result


def display_results(results):
    """显示分级抽奖结果"""
    print("\n" + "=" * 40)
    print("🎉 抽奖结果 🎉")
    for prize, winners in results.items():
        color_code = {
            "一等奖": "\033[1;31m",
            "二等奖": "\033[1;33m",
            "三等奖": "\033[1;32m"
        }[prize]

        print(f"\n{color_code}{prize}（{len(winners)}人）\033[0m")
        for i, winner in enumerate(winners, 1):
            print(f" 第{i}位：{color_code}{winner}\033[0m")
    print("\n" + "=" * 40)


def lottery_draw():
    """主抽奖程序"""
    participants = load_participants()
    if not participants:
        print("请检查participants.xlsx文件后重试")
        return

    results = multi_level_draw(participants)
    if results:
        display_results(results)


if __name__ == "__main__":
    print("=== 分级抽奖程序 ===")
    print("温馨提示：请确保当前目录存在participants.xlsx文件")
    print("           且包含名为'姓名'的列\n")

    while True:
        lottery_draw()
        choice = input("\n是否进行新一轮抽奖？(y/n) ").lower()
        if choice != "y":
            print("抽奖程序已结束！")
            break
