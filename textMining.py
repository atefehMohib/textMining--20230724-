# Import pandas and nltk libraries

import pandas as pd
import nltk
#nltk.download('punkt')

# Define the text data as a pandas series
text_data = pd.read_excel("text_mining.xlsx")

# Define a function to tokenize and normalize the text data
def tokenize_normalize(text):
    # Tokenize the text into words
    words = nltk.word_tokenize(text)
    # Return the list of words
    return words

text_data["text"].dropna(inplace=True)

# convert the column to string
text_data['text'] = text_data['text'].astype(str)

# Apply the function to the text data and create a new column
text_data["words"] = text_data["text"].apply(tokenize_normalize)
# Create a list of all words in the text data
all_words = []
for words in text_data["words"]:
    all_words.extend(words)
# Create a frequency distribution of all words
freq_dist = nltk.FreqDist(all_words)
# Find the top 10 most frequent words and their counts
top_10_words = freq_dist.most_common(10)
# Print the top 10 most frequent words and their counts
print("The most frequent words in the text data are:")
for word, count in top_10_words:
    print(word, count)

# Define a function to find the top three subjects for a given word
def find_top_three_subjects(word):
    # Create an empty list to store the subjects
    subjects = []
    # Loop through each row in the text data
    for index, row in text_data.iterrows():
    # Check if the word is in the row's words
        if word in row["words"]:
        # Loop through each word in the row's words
            for other_word in row["words"]:
            # Check if the other word is different from the word
                if other_word != word:
                # Append the other word to the subjects list
                    subjects.append(other_word)
                    # Create a frequency distribution of the subjects
                    subj_freq_dist = nltk.FreqDist(subjects)
                    # Find the top three most frequent subjects and their counts
                    top_three_subjects = subj_freq_dist.most_common(3)
                    # Return the top three most frequent subjects and their counts
    return top_three_subjects

# Print the top three subjects for each of the top 10 most frequent words
print("The top three subjects for each of the most frequent words are:")
for word, count in top_10_words:
    print(word, find_top_three_subjects(word))