import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """

st.markdown("""
<style>
.big-font {
    font-size:50px !important;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""<h1>Keyword Cannibalization Tool</h1> <p>Find pages that are competing with each other for the same 
keyword using first party GSC Data. The data must be multi-dimensional and contain query and page data. This can only be 
achieved using the GSC-API (exports from the GSC dashboard will not work).</p> <b>Directions: </b> <ul> <li>Upload query + page data (.csv) from the GSC API.</li> 
<li>Optional: Set a % threshold to select the top n% of queries (default 80%).</li> </ul>

</ul>
""", unsafe_allow_html=True)

gsc_data_file = st.file_uploader('Upload GSC Data', type='csv', key='key')
perc_slider = st.slider('Select % Threshold (ex: 80 = Selecting the top 80% of queries by metric)', 0, 100, value=80, step=10, key='int')


def is_ascii(string):
    # remove rows containing non-ascii characters
    try:
        string.encode('ascii')
    except UnicodeEncodeError:
        return False
    else:
        return True


def process_data(df, metric, perc_cumsum):
    # Group by the specified metric
    groupby_column_name = 'query'
    metric_name = metric
    grouped_df = df.groupby(metric, as_index=False).apply(lambda group: group.sort_values(metric_name, ascending=False))

    # Calculate sum per group and percentage per group
    sum_per_group = grouped_df.groupby(groupby_column_name)[metric_name].sum().sort_values(ascending=False)
    percent_per_group = grouped_df.groupby(groupby_column_name)[metric_name].sum() / grouped_df[metric_name].sum()

    # Calculate cumulative sum and get top 90% of groups
    cumsum_percent = percent_per_group.sort_values(ascending=False).cumsum()
    top_n_percent = percent_per_group[cumsum_percent <= float(perc_cumsum)]

    # Merge sum_per_group and top_90_percent back to the original DataFrame
    df = pd.merge(grouped_df, sum_per_group, left_on=groupby_column_name, right_index=True, suffixes=('', '_sum'))
    df = pd.merge(df, top_n_percent, left_on=groupby_column_name, right_index=True,
                  suffixes=('', '_percent_all_queries'))

    # Sort DataFrame
    df.sort_values([metric_name + '_sum', metric_name], ascending=[False, False], inplace=True)

    # Group by query and calculate the % total of each row within each group
    df['query_percentile_' + metric_name] = df[metric_name] / df[metric_name + '_sum']
    df.sort_values([metric_name + '_sum', metric_name], ascending=[False, False], inplace=True)

    # Keep rows where query_percentile is greater than or equal to 0.09
    df = df[df['query_percentile_' + metric_name] >= 0.10]

    # Keep only duplicate rows in the "query" column
    df = df[df.duplicated(subset=['query'], keep=False)]

    # Sort by metric_name + sum, then by metric_name, then by query
    df.sort_values([metric_name + '_sum', 'query', metric_name, 'position'], ascending=[False, True, False, True],
                   inplace=True)

    # Select specific columns
    df = df[[groupby_column_name, 'page', metric_name, 'ctr', 'position', metric_name + '_percent_all_queries',
             'query_percentile_' + metric_name]]
    return df


def process_merge(dfs):
    merged_df = pd.merge(dfs[0], dfs[1], on=['query', 'page'], how='inner')
    merged_df = merged_df[merged_df.duplicated(subset=['query'], keep=False)]
    merged_df.rename(columns={'ctr_x': 'ctr',
                              'position_x': 'position',
                              'query_perc_x': 'query_percent_clicks',
                              'query_perc_y': 'query_percent_impressions',
                              },
                     inplace=True)
    merged_df = merged_df[['query', 'page', 'clicks', 'impressions', 'ctr', 'position']]
    return merged_df


def format_excel(xlsx_file):
    wb = load_workbook(filename=xlsx_file)

    # Create a new objects for "Good" rows
    good_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    good_font = Font(color="006100")

    for ws in wb.worksheets:
        # Set zoom to 130%
        ws.sheet_view.zoomScale = 130

        # Keep track of queries already seen
        seen_queries = set()

        # Iterate through rows to find and highlight the first instance of each query
        for row in ws.iter_rows(min_row=2, max_col=ws.max_column, max_row=ws.max_row, values_only=False):
            query = row[0].value
            if query not in seen_queries:
                seen_queries.add(query)
                for cell in row:
                    cell.fill = good_fill
                    cell.font = good_font

        # Rename sheet
        ws.title = f'Competing by {ws.title}'

    # Save the modified workbook
    wb.save('cannibalization_analysis.xlsx')
    return wb


if __name__ == '__main__':
    if gsc_data_file is not None:
        data = pd.read_csv(gsc_data_file)
        data = data[data['query'].apply(lambda x: is_ascii(str(x)))]
        metrics = ['clicks', 'impressions', 'Impr. & Clicks']
        perc_cumsum = perc_slider / 100
        # remove rows with 0 clicks
        data = data[data['clicks'] > 0]
        dfs = []
        wb = Workbook()
        with st.spinner('Processing Data. . .'):
            for metric in metrics[:2]:
                df_processed = process_data(data, metric, perc_cumsum)
                dfs.append(df_processed)
        # merge dfs
        with st.spinner('Finalizing Data. . .'):
            dfs.append(process_merge(dfs))
        with st.spinner('Generating Spreadsheet. . .'):
            for sheet_name, df in zip(metrics, dfs):
                sheet = wb.create_sheet(title=sheet_name)
                # write dataframe to the sheet
                for row in dataframe_to_rows(df, index=False, header=True):
                    sheet.append(row)
            wb.remove(wb['Sheet'])
            # format excel
            wb.save('cannibalization_analysis.xlsx')
            with open("cannibalization_analysis.xlsx", "rb") as file:
                st.download_button(label='Download Cannibalization Analysis',
                                   data=file,
                                   file_name=f'cannibalization_analysis_threshold_{perc_cumsum}.xlsx',
                                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

st.write('---')
st.markdown(hide_streamlit_style, unsafe_allow_html=True)
st.write(
    'Author: [Tyler Gargula](https://tylergargula.dev) | Technical SEO & Software Developer | [Buy Me a Coffee](https://venmo.com/u/Tyler-Gargula)️☕️')
