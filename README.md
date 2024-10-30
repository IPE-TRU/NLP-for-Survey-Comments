# NLP-for-analyzing-survey-comments
This is a demo program for analyzing survey comments to identify and quantify key response themes. Prompts to get started with generating comments and code using AI is given below. Complete NLP analysis is available as code.

1. This is an example prompt to generate 100 sample comments for a survey question.

  Generate 100 different comments to the survey question ‘What additional products or services would you like to see in the future at the campus bookstore?’ and provide it as an excel file. To each comment add more columns for demographic information like this : "Age Gender Residency Campus Indigenous Faculty Program 29 W R K 0 Faculty of Law Juris Doctor Law"
  ![image](https://github.com/user-attachments/assets/55612cda-4486-4a15-83a6-d7d5eedbd83e)

2. This is an example prompt to get the code for NLP analysis of the survey responses.

Create a python program to analyse these student survey comments using natural language processing and NLTK package. Make sure to remove stopwords before analysis. The program should assign sentiment score and themes to each response. The themes are keywords from a dictionary with certain words mapped to certain keywords. The program should read input file and create three outputs. 1. Output one is a report in word. It should contain the the survey question, overall sentiment score for the responses, top 20 collocations, top 20 bigrams, top 10 trigrams, a wordcloud, and 5 sample comments wit demographic information for each comment. 2. The second output is an excel file with the response themes, number of comments and percentage of comments for each response theme, and percentage of comments per demographic breakdown such as resident, non resident, etc as different columns in the same sheet for each theme . 3. The third output is the comment list as an Excel file which contains each comment, themes mapped to that comment in the next column, and demographic columns columns in the same sheet. I will use this list to find if any comments are not assigned a theme. Then I will read the comment, decide the theme, and add it to the program until all comments are assigned relevant response themes. for example the comment "I would like to see more variety in snacks and beverages. And more space in aisles" can be mapped to two themes, {'Space': ['space', 'area', 'congest'], 'More food options': ['snack','beverages','food','drink','protein’].}.![image](https://github.com/user-attachments/assets/2fc2d199-f16e-4744-9bfc-0c09f3af7c2e)




