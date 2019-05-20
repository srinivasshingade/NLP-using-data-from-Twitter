import tweepy
import datetime
import xlsxwriter



# Twitter API credentials
consumer_api_key = "EnqVH1G1uUUGbWPWQT3Aouflq"
consumer_api_secret_key = "NTs23fJpq8UtsKCm6vyKTM6Cubyade1X6RaT6hVH2IcuSbHDuh"
access_token = "141925806-6HfOYWPKFe6FndCZQJ8JGzIBa3R0Hj1AI3TQZ6dT"
access_token_secret = "4pM5ioiSPC5BZEOWMY8LWaLtX86KXFVXtfDNykyA7hmns"



auth = tweepy.OAuthHandler(consumer_api_key, consumer_api_secret_key)
auth.set_access_token(access_token, access_token_secret)

api = tweepy.API(auth)


filter_words = ["deal", "deals", "blackfriday", "cybermonday" , "thanksgiving",  "sale"]


## '@' not required here
username_List = ["amazon","BestBuy","Walmart", "Target"]
startDate = datetime.datetime(2018, 11, 1, 0, 0, 0)
##endDate =   datetime.datetime.now()
##tweet.created_at < endDate and

tweetDictionary = {}

for username in username_List:
    print('\n\n' + "========  Fetching tweets for " + username + "  ========")
    tweets = []
     
    isNot_below_startDate = True
    fetchTweets = False
    print('\n' + "========  Fetching first 100 tweets for "+ username +"  ========")
    tmpTweets = api.user_timeline("@" + username, count = 100, tweet_mode="extended")
        
    while(isNot_below_startDate):
        if(fetchTweets):
            print("========  Fetching next 100 tweets for "+ username +"  ========")
            tmpTweets = api.user_timeline("@" + username, count = 100, max_id = tweets[-1].id - 1, tweet_mode="extended")
    
        oldestTweet_CreatedDate = tmpTweets[-1].created_at
        
        if(oldestTweet_CreatedDate > startDate):
            tweets.extend(tmpTweets);
        else:
             for tweet in tmpTweets:
                    if  tweet.created_at > startDate:
                        tweets.append(tweet)  
                        isNot_below_startDate = False
        fetchTweets = True 
        
    tweetDictionary[username] = tweets
    print('\n' + "========  Completed fetching tweets for "+ username +"  ========")
    print("========  Total tweets "+ str(len(tweets)) +"  ========")
    
    
    
    
tweetDictionary_ForProject = {}

print('\n\n' + "========  Writing tweets to Excel  ========")

workbook = xlsxwriter.Workbook("tweetDictionary.xlsx")

## Foreach KeyValuePair in Dictionary
## Key=Username; Value=List of Tweets obtained for Username

for username, tweets in tweetDictionary.items():
    worksheet = workbook.add_worksheet(username)
    
    tweet_Seller = []
    
    
    ## Table Header
    row = 0
    worksheet.write(row, 0, "Tweet_ID")
    worksheet.set_column(0, 0, 20)
    
    worksheet.write(row, 1, "Tweet_Timestamp")
    worksheet.set_column(1, 1, 20)
     
    worksheet.write(row, 2, "Tweet_Text")
    worksheet.set_column(2, 2, 85)
    
    worksheet.write(row, 3, "Is_Retweet_To")
    worksheet.set_column(3, 3, 20)
    
    worksheet.write(row, 4, "ReTweet_Count")
    worksheet.set_column(4, 4, 20)
    
    worksheet.write(row, 5, "Favourite_Count")
    worksheet.set_column(5, 5, 20)
    
    worksheet.write(row, 6, "Reply_Count")
    worksheet.set_column(6, 6, 20)
    
    worksheet.write(row, 7, "Place")
    worksheet.set_column(7, 7, 20)
    
    worksheet.write(row, 8, "Related_to_our_project")
    worksheet.set_column(8, 8, 23)
    row = 1    
    
    ## Loop Tweets
    for tweet in tweets:   
         worksheet.write_string(row, 0, str(tweet.id))
         worksheet.write_string(row, 1, str(tweet.created_at))
         worksheet.write(row, 2, tweet.full_text)
         worksheet.write_string(row, 3, str(tweet.in_reply_to_status_id))
         worksheet.write_string(row, 4, str(tweet.retweet_count))
         worksheet.write_string(row, 5, str(tweet.favorite_count))
         worksheet.write_string(row, 6, str(tweet.favorite_count))
         worksheet.write(row, 7, tweet.place)
         
        
         tweet_text = (tweet.full_text).lower().replace("!"," ").replace("?"," ").replace(","," ")
        
         if "thanks giving" in tweet_text:
            tweet_text = tweet_text.replace("thanks giving", "thanksgiving")        
         if "black friday" in tweet_text:
            tweet_text = tweet_text.replace("black friday", "blackfriday")
         if "cyber monday" in tweet_text:
            tweet_text = tweet_text.replace("cyber monday", "cybermonday")
            
         tweet_words = (tweet_text).split();
         intersection = [value for value in filter_words if value in tweet_words] 
        
         if len(intersection) > 0:
             worksheet.write(row, 8, "Yes")
             tweet_Seller.append(tweet)
             worksheet.set_row(row, options={'hidden': False})
         else:
             worksheet.write(row, 8, "No")
             worksheet.set_row(row, options={'hidden': True})
             
         row += 1
        
    worksheet.autofilter(0, 0, len(tweet_Seller), 8)
    worksheet.filter_column_list('I', ['Yes'])
    tweetDictionary_ForProject[username] = tweet_Seller
    

workbook.close()
print("Tweets Excel file ready")


print('\n\n' + "========  Getting Retweets  ========")
retweetDictionary = {}


for username, tweets in tweetDictionary_ForProject.items():
    
    retweets =[]
    print('\n\n' + "========  Getting Retweets for " + username +" ========")
    for tweet in tweets:
        if(tweet.retweet_count > 0):
            print("========  Getting Retweets for " + username + "for Tweet ID:"+str(tweet.id)+ " ========")
            retweets.extend(api.retweets(tweet.id, count=tweet.retweet_count))
    
         
    retweetDictionary[username] = retweets
    
    

print('\n\n' + "========  Writing ReTweets to Excel  ========")

workbook = xlsxwriter.Workbook("ReTweetDictionary.xlsx")

## Foreach KeyValuePair in Dictionary
## Key=Username; Value=List of Tweets obtained for Username

for username, tweets in retweetDictionary.items():
    worksheet = workbook.add_worksheet(username)
    
    tweet_Seller = []
    
    
    ## Table Header
    row = 0
    worksheet.write(row, 0, "Tweet_ID")
    worksheet.write(row, 1, "Tweet_Timestamp")
    worksheet.write(row, 2, "Tweet_Text")
    worksheet.write(row, 3, "Is_Retweet_To")
    worksheet.write(row, 4, "Favourite_Count")
    worksheet.write(row, 5, "Place")
    row = 1    
    
    ## Loop Tweets
    for tweet in tweets:   
         worksheet.write_string(row, 0, str(tweet.id))
         worksheet.write_string(row, 1, str(tweet.created_at))
         worksheet.write(row, 2, tweet.text)
         worksheet.write_string(row, 3, str(tweet.in_reply_to_status_id))
         worksheet.write_string(row, 4, str(tweet.favorite_count))
         worksheet.write(row, 5, tweet.place)
             
         row += 1    

workbook.close()
print("ReTweets Excel file ready")











