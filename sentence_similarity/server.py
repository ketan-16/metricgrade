import os
from pprint import pprint
import firebase_admin
from firebase_admin import credentials, firestore
from pprint import pprint


os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'key.json'
cred = credentials.ApplicationDefault()
firebase_admin.initialize_app(cred, {
    'Email-Classification-Service': 'projectId',
})
db = firestore.client()

test_dict = {}

docs = db.collection(u'drives').stream()
for doc in docs:
    try:
        student_attempts = doc.to_dict()
    except:
        'Attemps not present'
# for doc in docs:
#     pprint(f'{doc.id} => {doc.to_dict()}')
pprint(student_attempts)
