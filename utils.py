import re

import nltk
from nltk import download
from nltk.corpus import stopwords

download('stopwords', quiet=True)
download('punkt', quiet=True)

stop_words = stopwords.words('russian')


def stemmer(corpus):
    stem = nltk.SnowballStemmer("russian").stem
    stems = []
    for word in corpus:
        stems.append(stem(word))
    return stems


def tokenize(corpus):
    corpus = re.sub(r'[^\w\s]|_', ' ', corpus)  # замена скобок, пунктуации и "_" на " "

    tokens = [word for sent in nltk.sent_tokenize(corpus) for word in nltk.word_tokenize(sent)]
    valuable_words = []

    for token in tokens:
        token = token.lower().strip()
        if token.isalnum() and not token.isdigit() and token not in stop_words and stemmer([token])[0] not in stop_words:
            valuable_words.append(token)

    return valuable_words


def count_unique_words_with_registry(corpus, registry):
    result = 0
    for word in set(stemmer(tokenize(corpus))):
        if registry[word] == 1:
            result += 1
    return result


def count_unique_words_in_compare(corpus, row_to_compare_with):
    return len(set([stem for stem in stemmer(tokenize(corpus)) if stem not in stemmer(tokenize(row_to_compare_with))]))


def count_duplicates(corpus):
    stems = stemmer(tokenize(corpus))
    blacklist = []
    result = 0
    for index, stem in enumerate(stems[:-1]):
        if stem not in blacklist:
            result += len([token for token in stems[index+1:] if token in stem or stem in token])
            blacklist.append(stem)
    return result


def count_duplicates_between_rows(row1, row2):
    return len(set(stemmer(tokenize(row1))) & set(stemmer(tokenize(row2))))


def create_registry(data: list):
    registry = {}
    for row in data:
        for stem in set(stemmer(tokenize(row))):
            if stem in registry:
                registry[stem] += 1
            else:
                registry[stem] = 1

    return registry
