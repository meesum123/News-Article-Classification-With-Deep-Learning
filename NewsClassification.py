import xlwt
import tensorflow as tf
from sklearn import model_selection, preprocessing, linear_model, naive_bayes, metrics, svm
from sklearn.feature_extraction.text import TfidfVectorizer, CountVectorizer
from sklearn import decomposition, ensemble
import pandas, xgboost, numpy, textblob, string, nltk,textblob,keras
from keras.preprocessing import text, sequence
from keras import layers, models, optimizers
from keras.models import model_from_json
from keras.callbacks import ModelCheckpoint
config = tf.ConfigProto()
config.gpu_options.allow_growth = True
sess = tf.Session(config = config)

df = pandas.read_excel('C:\\Users\\meesum\\Downloads\\Participants_Data_News_category-20190729T063600Z-001\\Participants_Data_News_category\\Data_Train.xlsx', sheet_name='Sheet1')
trainDF = pandas.DataFrame()
dfv = pandas.read_excel('C:\\Users\\meesum\\Downloads\\Participants_Data_News_category-20190729T063600Z-001\\Participants_Data_News_category\\Data_Test.xlsx', sheet_name='Sheet1')
testDF = pandas.DataFrame()
testDF['text'] = dfv['STORY']
labels = df['SECTION']
trainDF['label'] = df['SECTION']
trainDF['text'] = df['STORY']
print('loaded dataset')

print(len(testDF['text']))

def train_model(classifier, feature_vector_train, label, feature_vector_valid, is_neural_net=False):
    # fit the training dataset on the classifier
    classifier.fit(feature_vector_train, label)
    
    # predict the labels on validation dataset
    predictions = classifier.predict(feature_vector_valid)
    
    if is_neural_net:
        predictions = predictions.argmax(axis=-1)
    
    return metrics.accuracy_score(predictions, valid_y)



def create_cnn():
    # Add an Input Layer
    input_layer = layers.Input((70, ))

    # Add the word embedding Layer
    embedding_layer = layers.Embedding(len(word_index) + 1, 300, weights=[embedding_matrix], trainable=False)(input_layer)
    embedding_layer = layers.SpatialDropout1D(0.3)(embedding_layer)

    # Add the convolutional Layer
    conv_layer = layers.Convolution1D(100, 3, activation="relu")(embedding_layer)

    # Add the pooling Layer
    pooling_layer = layers.GlobalMaxPool1D()(conv_layer)

    # Add the output Layers
    output_layer1 = layers.Dense(50, activation="relu")(pooling_layer)
    output_layer1 = layers.Dropout(0.25)(output_layer1)
    output_layer2 = layers.Dense(4, activation="softmax")(output_layer1)

    # Compile the model
    model = models.Model(inputs=input_layer, outputs=output_layer2)
    model.compile(optimizer = 'adam', loss = 'categorical_crossentropy', metrics = ['accuracy'])
    mc = ModelCheckpoint('best_model.h5', monitor='val_loss', mode='min', save_best_only=True)
    callbacks_list = [mc]
    model.fit(train_seq_x, train_y, batch_size=1000, epochs=500, validation_data=(valid_seq_x,valid_y), callbacks=callbacks_list)
    
    
    return model



c = 0

from nltk.stem import WordNetLemmatizer
lemmatizer = WordNetLemmatizer()
for article in trainDF['text']:
    word_list = nltk.word_tokenize(article)
    lemmatized_output = ' '.join([lemmatizer.lemmatize(w) for w in word_list])
    trainDF.at[c,'text'] = lemmatized_output
    c = c+1
c = 0
print('size of train')

print(len(trainDF['label']))
for article in testDF['text']:
    word_list = nltk.word_tokenize(article)
    lemmatized_output = ' '.join([lemmatizer.lemmatize(w) for w in word_list])
    testDF.at[c,'text'] = lemmatized_output
    c = c+1
print('size of test')
print(len(testDF['text']))
train_x, valid_x, train_y, valid_y = model_selection.train_test_split(trainDF['text'], trainDF['label'])
test_x = testDF['text'] 
print(len(train_x))
print(len(train_y))
print(len(valid_x))
print(len(valid_y))
print(len(test_x))

# label encode the target variable 

train_y = keras.utils.to_categorical(train_y)
valid_y = keras.utils.to_categorical(valid_y)

print("done encoding")
"""
tfidf_vect = TfidfVectorizer(analyzer='word', token_pattern=r'\w{1,}', max_features=5000)
tfidf_vect.fit(trainDF['text'])
xtrain_tfidf =  tfidf_vect.transform(train_x)
xvalid_tfidf =  tfidf_vect.transform(valid_x)
"""
# load the pre-trained word-embedding vectors 
import io

def load_vectors(fname):
    fin = io.open(fname, 'r', encoding='utf-8', newline='\n', errors='ignore')
    n, d = map(int, fin.readline().split())
    data = {}
    for line in fin:
        tokens = line.rstrip().split(' ')
        data[tokens[0]] = map(float, tokens[1:])
    return data
print("start embedding")
embeddings_index = {}
embeddings_index = load_vectors('C:\\Users\\meesum\\Downloads\\data\\wiki-news-300d-1M.vec')
# create a tokenizer 
token = text.Tokenizer()
token.fit_on_texts(pandas.concat([trainDF['text'],testDF['text']]))
word_index = token.word_index

# convert text to sequence of tokens and pad them to ensure equal length vectors 
train_seq_x = sequence.pad_sequences(token.texts_to_sequences(train_x), maxlen=70)
valid_seq_x = sequence.pad_sequences(token.texts_to_sequences(valid_x), maxlen=70)
test_seq_x = sequence.pad_sequences(token.texts_to_sequences(test_x), maxlen=70)
# create token-embedding mapping

embedding_matrix = numpy.zeros((len(word_index) + 1, 300))
for word, i in word_index.items():
    embedding_vector = embeddings_index.get(word)
    if embedding_vector is not None:
        embedding_matrix[i] = list(embedding_vector)

print("start training")
classifier = create_cnn()

#accuracy = train_model(classifier, train_seq_x, train_y, valid_seq_x, is_neural_net=True)
#print("CNN, Word Embeddings",  accuracy)
model_json = classifier.to_json()
with open("bestmodel.json", "w") as json_file:
    json_file.write(model_json)
print("saved weights and models to disk")
# serialize weights to HDF5

#classifier.save_weights("model.h5")
print("Saved model to disk")

json_file = open('bestmodel.json', 'r')
loaded_model_json = json_file.read()
json_file.close()
loaded_model = model_from_json(loaded_model_json)
# load weights into new model
loaded_model.load_weights("best_model.h5")
print("Loaded model from disk")

loaded_model.compile(optimizer = 'adam', loss = 'categorical_crossentropy', metrics = ['accuracy'])

print("Predictions")

workbook = xlwt.Workbook()  
  
sheet = workbook.add_sheet("Sub") 
  
# Specifying style 
style = xlwt.easyxf('font: bold 1') 
  
# Specifying column 
sheet.write(0, 0, 'SECTION', style) 
 
x = 1

y_pred = loaded_model.predict(test_seq_x)

for pred in y_pred:
    if(pred[0]>pred[1] and pred[0]>pred[2] and pred[0]>pred[3]):
        sheet.write(x, 0, 0)
        print(pred[0])
    elif(pred[1]>pred[0] and pred[1]>pred[2] and pred[1]>pred[3]):
        sheet.write(x, 0, 1)
        print(pred[1])
    elif(pred[2]>pred[1] and pred[2]>pred[0] and pred[2]>pred[3]):
        sheet.write(x, 0, 2)
        print(pred[2])
    elif(pred[3]>pred[1] and pred[3]>pred[2] and pred[3]>pred[0]):
        sheet.write(x, 0, 3)
        print(pred[3])
    else:
        sheet.write(x, 0, 1)
    x = x + 1

workbook.save("Submission.xls")

print('terminate plis')
print('terminate plis')
print('terminate plis')
print('terminate plis')
print('terminate plis')
print('terminate plis')
print('terminate plis')
print('terminate plis')
print('terminate plis')
print('terminate plis')
print('terminate plis')
print('terminate plis')
print('terminate plis')
print('terminate plis')
print('terminate plis')
