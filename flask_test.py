from flask import Flask

app = Flask(__name__)

@app.route('/')
def hello():
    return '✅ Flask is running!'

if __name__ == '__main__':
    print("✅ Flask API starting...")
    app.run(debug=True, port=5000)
