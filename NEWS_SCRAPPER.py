import streamlit as st
from newsapi import NewsApiClient
import json
import requests
from pandas.io.json import json_normalize
from datetime import datetime , timedelta
import pandas as pd


newsapi = NewsApiClient(api_key='d5f016a9f9be4a6d817211ca9833172b')

def top_headlines():
    today = datetime.now()
    start_date = today.strftime("%Y-%m-%d")
    end_date = (today + timedelta(days=7)).strftime("%Y-%m-%d")
    
  
    country = st.selectbox("Which country are you interested in?", ["in", "us", "gb", "de", "fr"])
    category = st.selectbox("Which category are you interested in?", ["business", "entertainment", "general", "health", "science", "technology"])
    keyword = st.text_input("Enter keyword")
    if st.button('Search'):
        news = newsapi.get_everything(q=keyword,
                                      from_param=start_date,
                                      to=end_date,
                                      language='en',
                                      category=category,
                                      country=country,
                                      sort_by='relevancy')
   
        
    if news["status"] == "ok":
        st.write("Results Found : ", news["totalResults"])
        df = pd.DataFrame(news['articles'])
        st.dataframe(df)
        all_articles = news["articles"] 
        for article in all_articles:
            st.write("**" + article["title"] + "**")
            st.write("Source : ", article["source"]["name"])
            st.write("Published at : ", article["publishedAt"])
            st.write(article["url"])
            st.write("\n\n")
       
    else:
        st.write("Sorry, No results found")

if __name__ == "__main__":
 st.title("Top Headlines")
 top_headlines()
