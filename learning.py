
from flask import Flask
print ("a")
app = Flask(__name__)

print ("abc")
@app.route('/')
def hello():
    return 'Hello, World!'
