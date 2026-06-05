# -*- coding: utf-8 -*-

import pandas as pd
import requests

INPUT_FILE = "organizer_dataset.xlsx"
OUTPUT_FILE = "organizer_classification_result.xlsx"
SUMMARY_FILE = "organizer_summary.txt"

MODEL = "llama3.2:3b"
SHEET_NAME = "Лист1"
LIMIT_ROWS = 100


def ask_llm(text: str) -> str:
    prompt = f"""
Ты классификатор обращений граждан.

Нужно определить, является ли обращение реальной проблемой.

Верни только один вариант:
проблема
не проблема

Правила:
1. Если есть жалоба на дороги, воду, мусор, отопление, транспорт, освещение, крышу, свалку, благоустройство, аварии, канализацию — это проблема.
2. Если это благодарность, вопрос, просьба сообщить телефон, график приема, куда обратиться, справочная информация — это не проблема.
3. Нельзя отвечать "не определено".
4. Нельзя объяснять ответ.

Текст обращения:
{text}

Ответ:
"""

    response = requests.post(
        "http://localhost:11434/api/generate",
        json={
            "model": MODEL,
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": 0
            }
        },
        timeout=120
    )

    answer = response.json()["response"].lower().strip()
    text_lower = text.lower()

    problem_keywords = [
        "без воды",
        "нет воды",
        "нет холодной воды",
        "нет горячей воды",
        "отключения воды",
        "ржавая вода",
        "давление воды",
        "нет отопления",
        "без отопления",
        "холодные батареи",
        "батареи",
        "мерзнут",
        "затопило",
        "затоплен",
        "подвал затоплен",
        "нечистотами",
        "канализация",
        "подтопление",
        "ямы",
        "яма",
        "разбит",
        "асфальт",
        "дорога не чищена",
        "дороги не чищены",
        "не чистят",
        "не чистилась",
        "снег не чистят",
        "почистите дорогу",
        "колеи",
        "гололед",
        "гололёд",
        "тротуар",
        "не убирают",
        "грязь",
        "мусор",
        "свалка",
        "помойка",
        "не вывозят",
        "не работает освещение",
        "нет освещения",
        "сломаны фонари",
        "крыша",
        "протекает",
        "автобус",
        "маршрут",
        "остановка",
        "опасно",
        "авария",
        "дтп",
        "интернета нет",
        "дым",
        "запах"
    ]

    non_problem_keywords = [
        "спасибо",
        "благодарю",
        "благодарность",
        "подскажите пожалуйста когда будут выплаты",
        "когда будут выплаты",
        "когда будет выплата",
        "когда будет пенсия",
        "выплаты на детей",
        "пособия за декабрь",
        "во сколько начало",
        "во сколько будут соревнования",
        "итоговый протокол",
        "какие документы",
        "как получить",
        "где можно получить",
        "график приема",
        "график приёма",
        "телефон ответственного",
        "сообщение без текста"
    ]

    for keyword in problem_keywords:
        if keyword in text_lower:
            return "проблема"

    for keyword in non_problem_keywords:
        if keyword in text_lower:
            return "не проблема"

    if "не проблема" in answer:
        return "не проблема"

    if "проблема" in answer:
        return "проблема"

    return "не проблема"


def main():
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)

    text_col = "Текст инцидента"
    municipality_col = "Муниципалитет"
    settlement_col = "Населенный пункт"
    group_col = "Группа тем"
    category_col = "Тема"

    sample = df.head(LIMIT_ROWS).copy()
    labels = []

    print("Файл:", INPUT_FILE)
    print("Всего строк в датасете:", len(df))
    print("Обрабатываем строк:", len(sample))
    print()
    print("Используемые колонки:")
    print("Текст:", text_col)
    print("Муниципалитет:", municipality_col)
    print("Населенный пункт:", settlement_col)
    print("Группа тем:", group_col)
    print("Тема:", category_col)
    print()

    for index, row in sample.iterrows():
        text = row[text_col]

        if pd.isna(text):
            labels.append("пусто")
            continue

        label = ask_llm(str(text))
        labels.append(label)
        print(f"{index}: {label} | {str(text)[:120]}")

    sample["LLM_классификация"] = labels
    problems = sample[sample["LLM_классификация"] == "проблема"]

    summary_by_municipality = (
        problems
        .groupby([municipality_col])
        .size()
        .reset_index(name="Количество проблем")
        .sort_values("Количество проблем", ascending=False)
    )

    summary_by_topic = (
        problems
        .groupby([municipality_col, group_col, category_col])
        .size()
        .reset_index(name="Количество")
        .sort_values("Количество", ascending=False)
    )

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        sample.to_excel(writer, sheet_name="classified_rows", index=False)
        summary_by_municipality.to_excel(writer, sheet_name="by_municipality", index=False)
        summary_by_topic.to_excel(writer, sheet_name="by_topic", index=False)

    with open(SUMMARY_FILE, "w", encoding="utf-8") as f:
        f.write("Краткая справка по тестовому датасету организаторов\n\n")
        f.write(f"Всего строк в датасете: {len(df)}\n")
        f.write(f"Обработано обращений: {len(sample)}\n")
        f.write(f"Выявлено проблем: {len(problems)}\n\n")

        f.write("Топ муниципалитетов по количеству проблем:\n")
        for _, row in summary_by_municipality.head(10).iterrows():
            f.write(f"- {row[municipality_col]} — {row['Количество проблем']} обращ.\n")

        f.write("\nТоп проблем по муниципалитетам и темам:\n")
        for _, row in summary_by_topic.head(10).iterrows():
            f.write(f"- {row[municipality_col]}: {row[group_col]} / {row[category_col]} — {row['Количество']} обращ.\n")

    print("\nГотово.")
    print(f"Excel-отчет: {OUTPUT_FILE}")
    print(f"Справка: {SUMMARY_FILE}")


if __name__ == "__main__":
    main()
