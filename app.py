# Complete content of app.py including sorted(nombres) on line 183 would be here.

# Sample code follows, replace with actual content:

from flask import Flask, jsonify

app = Flask(__name__)

nombres = ['Alice', 'Bob', 'Charlie']

@app.route('/nombres')
def get_nombres():
    return jsonify(sorted(nombres))

if __name__ == '__main__':
    app.run(debug=True)