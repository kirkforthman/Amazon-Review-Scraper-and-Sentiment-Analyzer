'''
Amazon Review Sentiment Analyzer
Kirk Forthman
December 26, 2023

A simple python script that scrapes Amazon product reviews, quantifies the sentiments contained in terms
of polarity [-1,1] and subjectivity [0,1] using TextBlob. The analysis of the reviews are then optionally
exported to an Excel file.
'''

import requests
import textblob
from bs4 import BeautifulSoup
import lxml
from textblob import TextBlob
import pandas as pd
import openpyxl


def get_data(url):
    r = requests.get(url, lxml)
    return r.text


def html_code(url):
    htmldata = get_data(url)
    soup = BeautifulSoup(htmldata, 'lxml')
    return soup


def polarity_to_text(text):
    polarity_value = TextBlob(text).sentiment.polarity
    if polarity_value >= 0.6:
        polarity_text = 'VERY POSITIVE'
    elif 0.2 <= polarity_value < 0.6:
        polarity_text = 'Positive'
    elif -0.2 <= polarity_value < 0.2:
        polarity_text = 'Neutral'
    elif -0.6 <= polarity_value < -0.2:
        polarity_text = 'Negative'
    else:
        polarity_text = 'VERY NEGATIVE'
    return polarity_text


def subjectivity_to_text(text):
    subj_value = TextBlob(text).sentiment.subjectivity
    if subj_value >= 0.8:
        subj_text = 'HIGHLY SUBJECTIVE'
    elif 0.6 <= subj_value < 0.8:
        subj_text = 'Somewhat Subjective'
    elif 0.4 <= subj_value < 0.6:
        subj_text = 'Neutral'
    elif 0.2 <= subj_value < 0.4:
        subj_text = 'Somewhat Objective'
    else:
        subj_text = 'VERY OBJECTIVE'
    return subj_text


if __name__ == '__main__':
    # switch out for lxml as needed for obfuscation needs
    HEADERS = ({'User-Agent':
                'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 '
                '(KHTML, like Gecko) Chrome/44.0.2403.157 Safari/537.36',
                'Accept-Language': 'en-US, en;q=0.5'})

    url = input('Please paste your Amazon product url here: ')
    soup = html_code(url)

    def reviews(soup):
        data_string = ''
        for item in soup.find_all('div', class_='a-row a-spacing-small review-data'):
            data_string = data_string + item.get_text()
        result = data_string.split('\n')
        return result

    review_data = reviews(soup)
    review_result = []

    for i in review_data:
        if i == '':
            pass
        else:
            review_result.append(i)

    # print results to terminal and collect aggregates
    review_number = 1
    invalid_language_count = 0
    raw_polarity = 0
    raw_subjectivity = 0
    review_number_list = []
    polarity_list = []
    subjectivity_list = []
    for review in review_result:
        if TextBlob(review).sentiment.polarity == 0 and TextBlob(review).sentiment.subjectivity == 0:
            print(f'Review #{review_number}: not written in supported language. No analysis performed.\n')
            invalid_language_count += 1
        else:
            print(f'Review #{review_number}: \n'
                  f'\tPolarity: {TextBlob(review).sentiment.polarity:.2f}; {polarity_to_text(review)}\n' 
                  f'\tSubjectivity: {TextBlob(review).sentiment.subjectivity:.2f}; {subjectivity_to_text(review)}\n')
            raw_polarity += TextBlob(review).sentiment.polarity
            raw_subjectivity += TextBlob(review).sentiment.subjectivity
        review_number_list.append(review_number)
        polarity_list.append(TextBlob(review).sentiment.polarity)
        subjectivity_list.append(TextBlob(review).sentiment.subjectivity)
        review_number += 1

    # aggregate results
    net_polarity = raw_polarity / (review_number - invalid_language_count)
    net_subjectivity = raw_subjectivity / (review_number - invalid_language_count - 1)
    print(f'\nPAGE TOTAL AGGREGATES:\n'
          f'\tThe average page values of {review_number - invalid_language_count - 1} reviews are:\n'
          f'\tAverage Polarity: {net_polarity}\n'
          f'\tAverage Subjectivity: {net_subjectivity}\n')

    # optional export data
    save_to_file = input('Save to File? (Y)es / (N)o:\n>')
    while save_to_file.lower() != 'y' and save_to_file.lower() != 'n':
        save_to_file = input('Invalid Input: Save to File? (Y)es / (N)o:\n>')
    if save_to_file.lower() == 'y':
        data = {'Polarity': polarity_list, 'Subjectivity': subjectivity_list}
        dataframe_out = pd.DataFrame(data=data)
        dataframe_out.to_excel('Example_Output.xlsx', sheet_name='SHEET_TITLE')
        print('File created.')
