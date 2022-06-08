from flask import Flask, render_template, request
from Excel_logic import *
app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def index():  # put application's code here
    excel = Excel('sample.xlsx')
    excel.get_data()
    return render_template('index.html', excel=excel)

if __name__ == '__main__':
    app.run()
