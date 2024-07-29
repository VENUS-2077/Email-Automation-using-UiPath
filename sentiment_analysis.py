import pandas as pd
from textblob import TextBlob

# Read data from file
file_path = r'D:\\Code\\Projects\\UiPath Projects\\EASA\\email_data.csv'
data = pd.read_csv(file_path)

# Basic sentiment classification and numeric sentiment scoring
def classify_sentiment(text):
    analysis = TextBlob(text)
    if analysis.sentiment.polarity > 0: # type: ignore
        return 'Positive', analysis.sentiment.polarity * 5 # type: ignore
    elif analysis.sentiment.polarity == 0: # type: ignore
        return 'Neutral', 0
    else:
        return 'Negative', analysis.sentiment.polarity * 5 # type: ignore

# Aspect-based sentiment analysis
def aspect_based_sentiment(text):
    aspects = {}
    # Simplified example, needs proper implementation
    sentences = text.split('.')
    for sentence in sentences:
        for word in sentence.split():
            if word.lower() in ['product', 'service', 'quality']:
                if word.lower() not in aspects:
                    aspects[word.lower()] = []
                aspects[word.lower()].append(TextBlob(sentence).sentiment.polarity) # type: ignore
    return aspects

# Analyze each email
results = []
for index, row in data.iterrows():
    sentiment, score = classify_sentiment(row['Body'])
    aspects = aspect_based_sentiment(row['Body'])
    results.append({
        'SenderName': row['SenderName'],
        'SenderEmail': row['SenderEmail'],
        'Subject': row['Subject'],
        'Date': row['Date'],
        'Body': row['Body'],
        'Sentiment': sentiment,
        'Score': score,
        'Aspects': aspects
    })

# Generate reports
summary = {'Positive': 0, 'Neutral': 0, 'Negative': 0}
for result in results:
    summary[result['Sentiment']] += 1

# Save results to files
results_df = pd.DataFrame(results)
results_df.to_excel(r'D:\\Code\\Projects\\UiPath Projects\\EASA\\sentiment_results.xlsx', index=False)
with open(r'D:\\Code\\Projects\\UiPath Projects\\EASA\\sentiment_report.txt', 'w') as file:
    file.write(f"Summary of Sentiment Analysis:\nPositive: {summary['Positive']}\nNeutral: {summary['Neutral']}\nNegative: {summary['Negative']}\n\n")
    file.write("Detailed Aspect-Based Sentiments:\n")
    for result in results:
        file.write(f"\nEmail from {result['SenderName']} ({result['SenderEmail']}):\n")
        file.write(f"Subject: {result['Subject']}\n")
        file.write(f"Date: {result['Date']}\n")
        file.write(f"Sentiment: {result['Sentiment']} (Score: {result['Score']})\n")
        file.write("Aspects:\n")
        for aspect, scores in result['Aspects'].items():
            file.write(f"  {aspect}: {sum(scores)/len(scores) if scores else 0}\n")
