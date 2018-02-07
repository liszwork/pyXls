# pyXls
openpyxl wrapper

# クラスメンバ

Xlsクラスのメンバ変数。

## book

開いたブック。

## sheet

開いたシート。

## data

loadAllData()で読み込んだデータが格納される。

## bookName

開いたブックの名称。

# メソッド

Xlsクラスのメソッド。

## __init__(self):

コンストラクタ。

インスタンス生成後、openBook()、openSheet()の実行が必要。

## __init__(self, bookName):

コンストラクタ。

インスタンス生成後、openSheet()の実行が必要。

引数で渡されたパスのブックをオープンする。

## __init__(self, bookName, sheetName):

コンストラクタ。

引数で渡されたパスのブックをオープンする。<br>
引数で渡されたシート名をオープンする。

## initialize(self):

現在未使用。

## openBook(self, bookName, isDataOnly=True):

引数で渡されたパスのブックを開く。

isDataOnlyにboolを渡すことで、openpyxl.load_workbook()のdata_onlyオプションを指定可能。<br>
指定しない場合、Trueが指定される。

## openSheet(self, sheetName):

引数で渡された名称のシートを開く。

## createSheet(self, sheetName, renew=False):

引数で渡された名称のシートを作成、または再作成する。

renewがTrueの場合で、同名のシートが既に存在したとき、シートを削除→作成する。

## resultText(self, result, text):

成功・失敗を付加したログを出力する。

resultの値に応じて、下記のように出力を行う。

> result=True: "[SUCCESS]" + text<br>
result=False: "[FAILED]" + text

## isOpenedBook(self, book, isOutputLog=False):

ブックが開いているかを判定する。

## isOpenedSheet(self, sheet, isOutputLog=False):

シートが開いているか確認する。

## isOpened(self, book, sheet, isOutputLog=False):

ブックとシートが開いているかを確認する。

## existSheet(self, sheetName):

現在のブックにシートが存在するかをチェックする。

存在する場合、Trueが返る。

## getCellValue(self, _row=0, _col=0):

引数で指定した行列番号のセルの値を取得する。

いずれかの値が0の場合、空文字が返る(失敗)。

## isBlankCell(self, val):

引数で渡した文字列が、有効な値かをチェックする。<br>
(取得したセルの値の有効チェック用)

## loadAllData(self, startRow=0, startCol=0):

指定した行列番号を始点とし、表のデータを取得する。

## writeHorizontal(self, data, row=1, col=1):

引数で渡されたデータを、指定した行列番号を始点として、右方向に順次書き込む。

データは配列を想定。

## save(self):

ブックを保存。

