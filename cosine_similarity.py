import numpy as np
import string
import nltk
import docx2txt
import os
import xlsxwriter as xw

folder= os.listdir(r"C:\Users\yunus\Desktop\All")

workbook = xw.Workbook('form_benzerlik_oranları.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write(0,0,"form1")
worksheet.write(0,1,"form2")
worksheet.write(0,2,"benzerlik")
a=0

for w in range (len(folder)):
    for q in range (w+1,len(folder)):
        
        try:
            sentence_1 = docx2txt.process(str(r"C:\Users\yunus\Desktop\All\\")+str(folder[w]))
            sentence_2 = docx2txt.process(str(r"C:\Users\yunus\Desktop\All\\")+str(folder[q]))
            
            def tokenizer(text):
                text_nopunct = "".join([char for char in text if char not in string.punctuation])
                tokens = nltk.word_tokenize(text_nopunct)
                return tokens
            
            word_list=list(set(tokenizer(sentence_1)+tokenizer(sentence_2)))
            
            vector_1 = [0]*len(word_list)
            vector_2 = [0]*len(word_list)
            
            def vectors(word_list, sentence, vector):
                for j in tokenizer(sentence):
                    for i in range(0,len(word_list)):
                        if j ==word_list[i]:
                            vector[i] =1
                vector = np.array(vector)
                return vector
                        
            vector_1 = vectors(word_list, sentence_1, vector_1)
            vector_2 = vectors(word_list, sentence_2, vector_2)
            
            
            def cos_sim(vector1, vector2):
                dot_product = np.dot(vector1, vector2)
                norm_1 = np.linalg.norm(vector1)
                norm_2 = np.linalg.norm(vector2)
                return dot_product / (norm_1 * norm_2)
            isim=str(folder[w])+" - "+str(folder[q])
            print(isim,cos_sim(vector_1,vector_2))
                        
            worksheet.write(a+1,0,str(folder[w]))
            worksheet.write(a+1,1,str(folder[q]))
            worksheet.write(a+1,2,cos_sim(vector_1,vector_2))
            
            a+=1
            
        except:
            continue
            
workbook.close()        























