import pandas as pd
import streamlit as st
from io import BytesIO
import re

# Function to process the Excel file
def process_excel(file):
    # Load all sheets from the Excel file
    excel_file = pd.ExcelFile(file)
    sheets = {sheet_name: excel_file.parse(sheet_name) for sheet_name in excel_file.sheet_names}

    # Access the required sheet 'Публикации'
    df = sheets['Публикации'].copy()

    # Step 1: Generate the list of unique companies from the 'Объект' column
    unique_companies = df['Объект'].dropna().unique().tolist()

    # Step 2: Filter out repeated news pieces for the same company/bank
    df = df.drop_duplicates(subset=['Объект', 'Выдержки из текста'])

    # Step 3: Define relevance assessment (including Russian-specific keywords)
    direct_keywords = ['убыток', 'прибыль', 'судебное дело', 'банкротство', 'потеря', 'миллиард', 'млн', 'миллиардов', 'миллионов', 'выручка']
    indirect_keywords = ['аналитик', 'комментарий', 'прогноз', 'отчет', 'заявление']

    def assess_relevance(text, company):
        if pd.isna(text):  # Handle missing values
            return 'unknown'
        
        text = text.lower()
        direct_relevance = any(keyword in text for keyword in direct_keywords) and company.lower() in text
        indirect_relevance = any(keyword in text for keyword in indirect_keywords)
        if direct_relevance and not indirect_relevance:
            return 'material'
        elif indirect_relevance:
            return 'not material'
        else:
            return 'unknown'

    # Step 3: Apply relevance assessment
    df.loc[:, 'Relevance'] = df.apply(lambda row: assess_relevance(row['Выдержки из текста'], row['Объект']), axis=1)

    # Step 4: Simple keyword-based sentiment assessment for Russian text
    negative_keywords = ['убыток', 'потеря', 'снижение', 'упадок', 'падение']
    positive_keywords = ['прибыль', 'рост', 'увеличение', 'подъем']

    def assess_sentiment(text):
        if pd.isna(text):  # Handle missing values
            return 'neutral'

        text = text.lower()
        if any(word in text for word in negative_keywords):
            return 'negative'
        elif any(word in text for word in positive_keywords):
            return 'positive'
        else:
            return 'neutral'

    df.loc[:, 'Sentiment'] = df['Выдержки из текста'].apply(assess_sentiment)

    # Step 5: Assess probable level of materiality based on financial amounts
    def assess_probable_materiality(text):
        if pd.isna(text):  # Handle missing values
            return 'unknown'

        match = re.search(r'(\d+)\s*млрд\s*руб', text.lower())
        if match:
            return 'significant'
        elif 'миллион' in text.lower():
            return 'significant'
        else:
            return 'not significant'

    df.loc[:, 'Materiality_Level'] = df['Выдержки из текста'].apply(assess_probable_materiality)

    # Step 6: Create Dashboard summarizing news for unique companies
    summary = df.groupby('Объект').agg(
        News_Count=('Выдержки из текста', 'count'),
        Risk_Level=('Materiality_Level', lambda x: 'high' if 'significant' in x.values else 'low')
    ).reindex(unique_companies, fill_value=0).reset_index()

    # Step 7: Filter only relevant news for the companies in the 'Объект' list
    filtered_news = df[df['Relevance'] == 'material']

    # Create a new Excel file with all sheets, adding 'Dashboard' and 'Filtered'
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write back the original sheets
        for sheet_name, data in sheets.items():
            data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Add modified sheets
        df.to_excel(writer, sheet_name='Публикации', index=False)
        summary.to_excel(writer, sheet_name='Dashboard', index=False)
        filtered_news.to_excel(writer, sheet_name='Filtered', index=False)

    output.seek(0)
    return output, filtered_news

# Streamlit app layout
st.title('Фильтр новостного файла в формате СКАН-Интерфакс на релевантность и значимость')
st.write("Загружайте и выгружайте!")

# File uploader
uploaded_file = st.file_uploader("Выбери Excel файл", type=["xlsx"])

if uploaded_file is not None:
    # Process the file when uploaded
    processed_file, filtered_table = process_excel(uploaded_file)

    # Display the 'Filtered' table on the web page
    st.write("Только материальные новости:")
    st.dataframe(filtered_table)

    # Download button for the processed file
    st.download_button(
        label="Ссылка на загрузку",
        data=processed_file,
        file_name="processed_news.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
