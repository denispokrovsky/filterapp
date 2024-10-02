import streamlit as st
import pandas as pd
import re
from io import BytesIO
from rapidfuzz import fuzz
from openpyxl import load_workbook
from langchain.llms import OpenAI
from langchain.chains import LLMChain
from langchain.prompts import PromptTemplate

# Streamlit app layout
st.set_page_config(page_title="::: мониторинг новостного потока :::", layout="wide")

st.title('Фильтр новостного файла в формате СКАН-Интерфакс на релевантность и значимость!')
st.write("Загружайте и выгружайте!")

# File uploader
uploaded_file = st.file_uploader("Выбери Excel файл", type=["xlsx"])

# Access the OpenAI API key from Streamlit secrets
openai_api_key = st.secrets["OPENAI_API_KEY"]

# Initialize OpenAI LLM
llm = OpenAI(model_name="gpt-4o-mini", temperature=0.7, openai_api_key=openai_api_key)

# Define prompt templates for LangChain
risk_prompt_template = PromptTemplate(
    input_variables=["text", "company"],
    template="Текст: {text}\nКомпания: {company}\nЕсть ли риск убытка для этой компании? Ответьте 'Риск убытка' или 'Нет риска убытка'."
)

comment_prompt_template = PromptTemplate(
    input_variables=["text", "company"],
    template="Текст: {text}\nКомпания: {company}\nСколько примерно может потерять компания? Дайте комментарий не более 200 слов."
)

# Function to process the Excel file and add new columns
def process_excel_with_fuzzy_matching(file, sample_file, similarity_threshold=65):
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
    def fuzzy_deduplicate(df, column, threshold=65):
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

    # Step 6: Risk assessment using LangChain
    risk_chain = LLMChain(llm=llm, prompt=risk_prompt_template)
    comment_chain = LLMChain(llm=llm, prompt=comment_prompt_template)

    # Using apply() to handle multiple inputs (text and company)
    df_deduplicated['Risk of loss'] = df_deduplicated.apply(lambda row: risk_chain.apply([{"text": row['Выдержки из текста'], "company": row['Объект']}])[0], axis=1)
    df_deduplicated['Comment'] = df_deduplicated.apply(lambda row: comment_chain.apply([{"text": row['Выдержки из текста'], "company": row['Объект']}])[0], axis=1)

    # Step 7: Prepare summary for "Сводка" sheet
    dashboard_summary = df_deduplicated.groupby('Объект').agg(
        News_Count=('Выдержки из текста', 'count'),
        Significant_Texts=('Materiality_Level', lambda x: (x == 'значительна').sum()),
        Negative_Texts=('Sentiment', lambda x: (x == 'негатив').sum()),
        Positive_Texts=('Sentiment', lambda x: (x == 'позитив').sum()),
        Risk_Level=('Materiality_Level', lambda x: 'высокий' if 'значительна' in x.values else 'низкий')
    ).reset_index()

    # Sort the summary by News_Count first and Significant_Texts second (both in descending order)
    dashboard_summary_sorted = dashboard_summary.sort_values(by=['Significant_Texts', 'News_Count'], ascending=[False, False])

    # Create new dashboard summary filtered by 'Риск убытка'
    new_dashboard_summary = df_deduplicated[df_deduplicated['Risk of loss'] == 'Риск убытка'][['Объект', 'Заголовок', 'Выдержки из текста', 'Risk of loss', 'Comment']]

    # Rename columns for display in Streamlit
    dashboard_summary_sorted.columns = [
        'Компания',
        'Всего публикаций',
        'Из них: материальных',
        'Из них: негативных',
        'Из них: позитивных',
        'Уровень материального негатива'
        ]

    filtered_news = df_deduplicated[df_deduplicated['Relevance'] == 'материальна']
    filtered_news = filtered_news.drop_duplicates(subset=['Объект', 'Выдержки из текста']).reset_index(drop=True)

    # Load the sample Excel file to maintain formatting
    book = load_workbook(sample_file)

    # Write sorted data to the "Сводка" sheet
    dashboard_sheet = book['Сводка']
    for idx, row in dashboard_summary_sorted.iterrows():
        dashboard_sheet[f'E{4 + idx}'] = row['Компания']
        dashboard_sheet[f'F{4 + idx}'] = row['Всего публикаций']
        dashboard_sheet[f'G{4 + idx}'] = row['Из них: материальных']
        dashboard_sheet[f'H{4 + idx}'] = row['Из них: негативных']
        dashboard_sheet[f'I{4 + idx}'] = row['Из них: позитивных']
        dashboard_sheet[f'J{4 + idx}'] = row['Уровень материального негатива']

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

    return output, df_deduplicated, original_news_count, duplicates_removed, remaining_news_count, dashboard_summary_sorted, new_dashboard_summary


# Handle file upload and processing
if uploaded_file is not None:
    # Store the path to the sample Excel file for formatting
    sample_file = "sample_file.xlsx"

    # Process the file and get the processed output, filtered data, and counts
    processed_file, filtered_table, original_news_count, duplicates_removed, remaining_news_count, dashboard_summary_sorted, new_dashboard_summary = process_excel_with_fuzzy_matching(uploaded_file, sample_file)

    # Display the filtered news as it appears in Excel
    st.write(f"Из {original_news_count} новостных сообщений удалены {duplicates_removed} дублирующих. Осталось {remaining_news_count}.")
    
    st.write("Только материальные новости:")
    st.dataframe(filtered_table[['Объект', 'Relevance', 'Sentiment', 'Materiality_Level', 'Заголовок', 'Выдержки из текста']])

    # Display the sorted dashboard summary
    st.write("Сводка:")
    st.dataframe(dashboard_summary_sorted)

    # Display the new dashboard summary filtered by the presence of 'Риск убытка'
    st.write("Сводка с риском убытка:")
    st.dataframe(new_dashboard_summary)

    # Provide a download button for the processed file
    st.download_button(
        label="СКАЧАЙ ЗДЕСЬ:",
        data=processed_file,
        file_name="processed_news.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
