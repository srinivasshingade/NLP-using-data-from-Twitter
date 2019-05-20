import pandas as pd
import nltk
from nltk.corpus import stopwords
import string
import gensim
from gensim import corpora
import datetime
import os


os.chdir(r'C:\Users\A1814203\Downloads\670\Final Project')

cwd = os.getcwd()
data_Directory = cwd + "\Data"
results_Directory = cwd + "\Results"
xl_FilePath = cwd + "\Data\Dataset_TopicModeling.xlsx"

company_list = ["amazon", "BestBuy","Target", "Walmart"]

fullData = {}
tweetsDictionary = {}

# Load spreadsheet
xl = pd.ExcelFile(xl_FilePath)

for company in company_list:
    fullData[company] = xl.parse(company, header=[0])
    df = fullData[company]
    #df_UsefulForProject = df[df['Related_to_our_project'] == 'Yes']
    tweetsDictionary[company] = df["Tweet_Text"]


#Stop Words
#filter_words = ["deal", "deals", "blackfriday", "cybermonday" , "thanksgiving",  "sale"]
stop_words = stopwords.words("english")
#stop_words.extend(filter_words)
stops = set(stop_words)

stemmer = nltk.stem.SnowballStemmer('english')
lemmatizer = nltk.wordnet.WordNetLemmatizer()

#LDA - Preprocessing for TM
processedTweetsDictionary = {}

for company, tweets in tweetsDictionary.items():
    processedTweets = []
    for tweet in tweets:
        
#        tweet = tweet.lower()
#        if "thanks giving" in tweet:
#            tweet = tweet.replace("thanks giving", "thanksgiving")        
#        if "black friday" in tweet:
#            tweet = tweet.replace("black friday", "blackfriday")
#        if "cyber monday" in tweet:
#            tweet = tweet.replace("cyber monday", "cybermonday")
        
        processedTweet = tweet.strip()
        processedTweet = processedTweet.translate(str.maketrans('','',string.punctuation))
        tweetTokens = nltk.word_tokenize(processedTweet)
        
        tweetTokens = list(set(tweetTokens) - stops)
        
        for token in tweetTokens:
            oldToken = token
            tweetTokens.remove(token)
            oldToken = lemmatizer.lemmatize(oldToken)
            tweetTokens.append(oldToken)
        
        processedTweet = ' '.join(tweetTokens)
        print(processedTweet)
        
        if len(tweetTokens) > 3:
            processedTweets.append(processedTweet)
      
    processedTweetsDictionary[company] = processedTweets
    
    
    
#LDA - Topic Modeling   
ldaModel_Dictionary = {} 
ldaReport_Dictionary = {}
topicCount = 10
passCount = 1
start = datetime.datetime.now()

     
for company, processedTweets in processedTweetsDictionary.items():
    
    result = "\n\n==> Topic Modeling for " + company + ":"
    result += "\n    Total Tweets used for Topic Modeling : " + str(len(processedTweets))
    print(result)
    
    company_start = datetime.datetime.now()
    
    texts = [[text for text in doc.split()] for doc in processedTweets]
    print(texts)
    dictionary = corpora.Dictionary(texts)
    #print("printing dictionary",dictionary.token2id)
    #print(dictionary)
    doc_term_matrix = [dictionary.doc2bow(doc.split()) for doc in processedTweets]
    #print(doc_term_matrix)
    ldaObject = gensim.models.ldamodel.LdaModel
    ldaModel = ldaObject(doc_term_matrix, num_topics=topicCount, id2word=dictionary,passes=passCount)
    company_end = datetime.datetime.now()
    
    ldaResult = str(ldaModel.print_topics(num_topics=topicCount, num_words=10))
    result += "\n    Start Time : " + str(company_start) + "\t\t End Time : " + str(company_end)  + "\t\t Time elapsed: " + str(company_end - company_start)
    result += "\n\nResult:\n" + ldaResult + "\n\n"
    
    ldaModel_Dictionary[company] = ldaModel
    ldaReport_Dictionary[company] = result
    
    
    
    print(result)

end = datetime.datetime.now()


finalString = "Topic Modeling for tweets by companies below\n" + ', '.join(company_list)
finalString += "\n\nTopics : " + str(topicCount) + "\tPasses : " + str(passCount) 
finalString += "\nStart Time : " + str(start) + "\t\t End Time : " + str(end)  + "\t\t Time elapsed: " + str(end - start) + "\n"
for company, result in ldaReport_Dictionary.items():
    finalString += "\n" + ("==== " * 45)
    finalString += "\n" + str(result.replace("), ", "),\n"))
    finalString += "\n" + ("==== " * 45)
    
    
if not os.path.exists(results_Directory):
    os.makedirs(results_Directory)

file = open(results_Directory + "\Result_" + end.strftime("%Y%m%d-%H%M%S") + ".txt", "w", encoding='utf8') 
file.write(finalString) 
file.close() 

print("\n\nLDA analysis complete")


#LDA Visulaization
import pyLDAvis.gensim
for company, processedTweets in processedTweetsDictionary.items():
    
    texts = [[text for text in doc.split()] for doc in processedTweets]
    dictionary = corpora.Dictionary(texts)
    doc_term_matrix = [dictionary.doc2bow(doc.split()) for doc in processedTweets]
    
    lda_display = pyLDAvis.gensim.prepare(ldaModel_Dictionary[company], doc_term_matrix, dictionary, sort_topics=False)
    pyLDAvis.save_html(lda_display, results_Directory + '\\' +company + '_lda.html')











    
