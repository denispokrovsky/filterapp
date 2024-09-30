import streamlit as st
import pandas as pd
import re
from io import BytesIO
from rapidfuzz import fuzz
from openpyxl import load_workbook

# Streamlit app layout
st.title('Фильтр новостного файла в формате СКАН-Интерфакс на релевантность и значимость!')
st.write("Загружайте и выгружайте!")

# File uploader
uploaded_file = st.file_uploader("Выбери Excel файл", type=["xlsx"])

def process_excel_with_fuzzy_matching(file, sample_file, similarity_threshold=70):
    # Load all sheets from the uploaded Excel file
    excel_file = pd.ExcelFile(file)
    sheets = {sheet_name: excel_file.parse(sheet_name) for sheet_name in excel_file.sheet_names}

    # Access the required sheet 'Публикации'
    df = sheets['Публикации'].copy()

    # Track original number of news (X)
    original_news_count = len(df)

    # Step 1: Generate the list of unique companies from the 'Объект' column
    unique_companies = df['Объект'].dropna().unique().tolist()

    # Step 2: Fuzzy filter out similar news for the same company/bank
    def fuzzy_deduplicate(df, column, threshold=90):
        seen_texts = []
        indices_to_keep = []
        for i, text in enumerate(df[column]):
            if pd.isna(text):
                indices_to_keep.append(i)
                continue
            text = str(text)
            if not seen_texts or all(fuzz.ratio(text, seen) < threshold for seen in seen_texts):
                seen_texts.append(text)
                indices_to_keep.append(i)
        return df.iloc[indices_to_keep]

    # Apply fuzzy deduplication on 'Выдержки из текста' column
    df_deduplicated = df.groupby('Объект').apply(lambda x: fuzzy_deduplicate(x, 'Выдержки из текста', similarity_threshold)).reset_index(drop=True)

    # Track the number of remaining news (Z)
    remaining_news_count = len(df_deduplicated)

    # Calculate number of duplicates removed (Y)
    duplicates_removed = original_news_count - remaining_news_count

    # Step 3: Define relevance assessment (including Russian-specific keywords)
    direct_keywords = ['убыток', 'прибыль', 'судебное дело', 'банкротство', 'потеря', 'миллиард', 'млн', 'миллиардов', 'миллионов', 'выручка']
    indirect_keywords = ['аналитик', 'комментарий', 'прогноз', 'отчет', 'заявление']

    def assess_relevance(text, company):
        if pd.isna(text):
            return 'н/д'
        text = text.lower()
        direct_relevance = any(keyword in text for keyword in direct_keywords) and company.lower() in text
        indirect_relevance = any(keyword in text for keyword in indirect_keywords)
        if direct_relevance and not indirect_relevance:
            return 'материальна'
        elif indirect_relevance:
            return 'нематериальная'
        else:
            return 'н/д'

    df_deduplicated['Relevance'] = df_deduplicated.apply(lambda row: assess_relevance(row['Выдержки из текста'], row['Объект']), axis=1)

    # Step 4: Sentiment assessment
    negative_keywords = ['убыток', 'потеря', 'снижение', 'упадок', 'падение']
    positive_keywords = ['прибыль', 'рост', 'увеличение', 'подъем']

    def assess_sentiment(text):
        if pd.isna(text):
            return 'нейтрально'
        text = text.lower()
        if any(word in text for word in negative_keywords):
            return 'негатив'
        elif any(word in text for word in positive_keywords):
            return 'позитив'
        else:
            return 'нейтрально'

    df_deduplicated['Sentiment'] = df_deduplicated['Выдержки из текста'].apply(assess_sentiment)

    # Step 5: Materiality assessment
    def assess_probable_materiality(text):
        if pd.isna(text):
            return 'н/д'
        match = re.search(r'(\d+)\s*млрд\s*руб', text.lower())
        if match:
            return 'значительна'
        elif 'миллион' in text.lower():
            return 'значительна'
        else:
            return 'незначительна'

    df_deduplicated['Materiality_Level'] = df_deduplicated['Выдержки из текста'].apply(assess_probable_materiality)

    # Step 6: Prepare summary for "Сводка" sheet
    dashboard_summary = df_deduplicated.groupby('Объект').agg(
        News_Count=('Выдержки из текста', 'count'),
        Significant_Texts=('Materiality_Level', lambda x: (x == 'значительна').sum()),
        Negative_Texts=('Sentiment', lambda x: (x == 'негатив').sum()),
        Positive_Texts=('Sentiment', lambda x: (x == 'позитив').sum()),
        Risk_Level=('Materiality_Level', lambda x: 'высокий' if 'значительна' in x.values else 'низкий')
    ).reset_index()

    # Sort the summary by News_Count first and Significant_Texts second (both in descending order)
    dashboard_summary = dashboard_summary.sort_values(by=['News_Count', 'Significant_Texts'], ascending=[True, True])

    # Step 7: Filter only material news, ensuring non-duplicate texts
    filtered_news = df_deduplicated[df_deduplicated['Relevance'] == 'материальна']
    filtered_news = filtered_news.drop_duplicates(subset=['Объект', 'Выдержки из текста']).reset_index(drop=True)

    # Load the sample Excel file to maintain formatting
    book = load_workbook(sample_file)

    # Write to the "Сводка" sheet
    dashboard_sheet = book['Сводка']
    for idx, row in dashboard_summary.iterrows():
        dashboard_sheet[f'E{4 + idx}'] = row['Объект']
        dashboard_sheet[f'F{4 + idx}'] = row['News_Count']
        dashboard_sheet[f'G{4 + idx}'] = row['Significant_Texts']
        dashboard_sheet[f'H{4 + idx}'] = row['Negative_Texts']
        dashboard_sheet[f'I{4 + idx}'] = row['Positive_Texts']
        dashboard_sheet[f'J{4 + idx}'] = row['Risk_Level']

    # Write to the 'Публикации' sheet
    publications_sheet = book['Публикации']
    for r_idx, row in df_deduplicated.iterrows():
        for c_idx, value in enumerate(row):
            publications_sheet.cell(row=2 + r_idx, column=c_idx + 1).value = value

    # Write to the 'Значимые' sheet, no empty rows
    filtered_sheet = book['Значимые']
    for f_idx, row in filtered_news.iterrows():
        filtered_sheet[f'C{3 + f_idx}'] = row['Объект']
        filtered_sheet[f'D{3 + f_idx}'] = row['Relevance']
        filtered_sheet[f'E{3 + f_idx}'] = row['Sentiment']
        filtered_sheet[f'F{3 + f_idx}'] = row['Materiality_Level']
        filtered_sheet[f'G{3 + f_idx}'] = row['Заголовок'] if 'Заголовок' in row else ''
        filtered_sheet[f'H{3 + f_idx}'] = row['Выдержки из текста']

    # Save the final file to a BytesIO buffer
    output = BytesIO()
    book.save(output)
    output.seek(0)

    return output, filtered_news, original_news_count, duplicates_removed, remaining_news_count

# Handle file upload and processing
if uploaded_file is not None:
    # Store the path to the sample Excel file for formatting
    sample_file = "sample_file.xlsx"

    # Process the file and get the processed output, filtered data, and counts
    processed_file, filtered_table, original_news_count, duplicates_removed, remaining_news_count = process_excel_with_fuzzy_matching(uploaded_file, sample_file)

    # Display the filtered news as it appears in Excel
    st.write(f"Из {original_news_count} новостных сообщений удалены {duplicates_removed} дублирующих. Осталось {remaining_news_count}.")
    
    st.write("Только материальные новости:")
    st.dataframe(filtered_table[['Объект', 'Relevance', 'Sentiment', 'Materiality_Level', 'Заголовок', 'Выдержки из текста']])

    # Provide a download button for the processed file
    st.download_button(
        label="СКАЧАЙ ЗДЕСЬ:",
        data=processed_file,
        file_name="processed_news.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
