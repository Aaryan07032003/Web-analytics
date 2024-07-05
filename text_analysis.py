import requests
from bs4 import BeautifulSoup
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
import pandas as pd
import re

nltk.download('punkt')
nltk.download('stopwords')

stopwords_files = [
    "StopWords_Auditor.txt", "StopWords_Currencies.txt", "StopWords_DatesandNumbers.txt",
    "StopWords_Generic.txt", "StopWords_GenericLong.txt", "StopWords_Geographic.txt",
    "Stopwords_Names.txt"
]

stopwords_set = set()
for file in stopwords_files:
    with open(file, 'r') as f:
        stopwords_set.update(f.read().splitlines())

with open("positive-words.txt", 'r') as f:
    positive_words = set(f.read().splitlines())

with open("negative-words.txt", 'r') as f:
    negative_words = set(f.read().splitlines())

def extract_text_from_url(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')

        title = soup.find('title').get_text()
        paragraphs = soup.find_all('p')
        text = " ".join([p.get_text() for p in paragraphs])

        return title, text
    except Exception as e:
        print(f"Error extracting {url}: {e}")
        return None, None

def clean_text(text):
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[^a-zA-Z\s]', '', text)
    text = text.lower()
    tokens = word_tokenize(text)
    cleaned_text = [word for word in tokens if word not in stopwords_set]
    return cleaned_text

def analyze_text(text):
    sentences = sent_tokenize(text)
    words = word_tokenize(text)
    cleaned_words = clean_text(text)

    positive_score = sum(1 for word in cleaned_words if word in positive_words)
    negative_score = sum(1 for word in cleaned_words if word in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(cleaned_words) + 0.000001)

    avg_sentence_length = len(words) / len(sentences)
    complex_words_count = sum(1 for word in cleaned_words if syllable_count(word) > 2)
    percentage_complex_words = complex_words_count / len(cleaned_words)
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

    avg_words_per_sentence = len(words) / len(sentences)
    syllables_per_word = sum(syllable_count(word) for word in cleaned_words) / len(cleaned_words)
    personal_pronouns = sum(1 for word in words if re.match(r'\b(I|we|my|ours|us)\b', word))
    avg_word_length = sum(len(word) for word in cleaned_words) / len(cleaned_words)

    return {
        "POSITIVE SCORE": positive_score,
        "NEGATIVE SCORE": negative_score,
        "POLARITY SCORE": polarity_score,
        "SUBJECTIVITY SCORE": subjectivity_score,
        "AVG SENTENCE LENGTH": avg_sentence_length,
        "PERCENTAGE OF COMPLEX WORDS": percentage_complex_words,
        "FOG INDEX": fog_index,
        "AVG NUMBER OF WORDS PER SENTENCE": avg_words_per_sentence,
        "COMPLEX WORD COUNT": complex_words_count,
        "WORD COUNT": len(cleaned_words),
        "SYLLABLE PER WORD": syllables_per_word,
        "PERSONAL PRONOUNS": personal_pronouns,
        "AVG WORD LENGTH": avg_word_length
    }

def syllable_count(word):
    word = word.lower()
    count = len(re.findall(r'[aeiouy]', word))
    if word.endswith('es') or word.endswith('ed'):
        count -= 1
    return max(count, 1)

input_file = 'Input.xlsx'
output_file = 'Output Data Structuress.xlsx'

input_df = pd.read_excel(input_file)

results = []

for index, row in input_df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']

    print(f"Processing URL ID {url_id}: {url}")
    title, text = extract_text_from_url(url)

    if text:
        analysis_results = analyze_text(text)
        analysis_results['URL_ID'] = url_id
        analysis_results['URL'] = url
        results.append(analysis_results)

columns = [
    "URL_ID", "URL", "POSITIVE SCORE", "NEGATIVE SCORE", "POLARITY SCORE", "SUBJECTIVITY SCORE",
    "AVG SENTENCE LENGTH", "PERCENTAGE OF COMPLEX WORDS", "FOG INDEX", "AVG NUMBER OF WORDS PER SENTENCE",
    "COMPLEX WORD COUNT", "WORD COUNT", "SYLLABLE PER WORD", "PERSONAL PRONOUNS", "AVG WORD LENGTH"
]

output_df = pd.DataFrame(results, columns=columns)
output_df.to_excel(output_file, index=False)

print(f"Analysis complete. Results saved to {output_file}")
