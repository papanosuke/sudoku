# -*- encoding: utf-8 -*-

require 'win32ole'
#*************************************************************
#*サブルーチン
#*************************************************************
#*****************
#*フルパス取得
#*****************
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end
#*****************
#*Fileオープン
#*****************
def openExcelWorkbook filename
  filename = getAbsolutePath(filename)
  xl = WIN32OLE.new('Excel.Application')
  xl.Visible = false
  xl.DisplayAlerts = false
  book = xl.Workbooks.Open(filename)
  begin
    yield book
  ensure
    xl.Workbooks.Close
    xl.Quit
  end
end
#*****************
#*file作成
#*****************
def createExcelWorkbook
  xl = WIN32OLE.new('Excel.Application')
  xl.Visible = false
  xl.DisplayAlerts = false
  book = xl.Workbooks.Add()
  begin
    yield book
  ensure
    xl.Workbooks.Close
    xl.Quit
  end
end

#*************************************************************
#*メインルーチン
#*************************************************************
=begin
問題作成の考え方
①ランダムで1～9の数字を設定する
②縦、横、ボックスの３通りのチェックを実施する
③①～②の繰り返し

=end

ERR_OK = 0
ERR_NG = 1

#*************************************************************
#*クラス
#*************************************************************
class Sudoku
  attr_accessor :sdkarea
  attr_accessor :st_x
  attr_accessor :ed_x
  attr_accessor :st_y
  attr_accessor :ed_y
  
#********************
#*三次配列
#********************
  def thirdarray(i,j,k)
    (0...i).map {
      (0...j).map {
       Array.new(k)
      }
    }
  end
#
#********************
#*二次配列
#********************
  def secondarray(i,j)
    (0...i).map {
       Array.new(j)
    }
  end
#
#********************
#*一次配列
#********************
  def firstarray(i)
    Array.new(i)
  end
#
#********************
#*初期化
#********************
  def initialize
    @sdkarea = self.secondarray(9,9)
    for i in 0..8 do
      for j in 0..8 do
        @sdkarea[i][j] = 0
      end
    end

    @st_x = 0
    @st_y = 0
    @ed_x = 0
    @ed_y = 0

  end
#
#********************
#エラーチェック
#********************
  def errchk(sdkarea, st_x, st_y, ed_x ,ed_y)

    chkarea = []
    for i in 0..8 do
      chkarea[i] = 0
    end

    for i in st_x..ed_x do
      for j in st_y..ed_y do
        
#選択された数字の重複をチェック
        if sdkarea[i][j] != 0
          chkarea[sdkarea[i][j] - 1] += 1

        end

#       print " s=",sdkarea[i],chkarea,"\n"
#
#重複した場合エラーを返す
        return ERR_NG if chkarea[sdkarea[i][j] - 1] > 1
      end

    end

    
    return ERR_OK
  end
#
end

=begin
#*************************************************************
#*メインルーチン
#*************************************************************
=end
kekka = Sudoku::new

#*************
#問題作成
#*************
i = 0
j = 0
while true
  while true
#
    #数字をランダムで生成
    kekka.sdkarea[i][j] = rand(9)+1
#
#===>DEBUG START
#   print "i=",i," j=",j," suji=",kekka.sdkarea[i][j],"\n"
#<===DEBUG END
=begin
(0,0)	(0,1)	(0,2)	(0,3)	(0,4)	(0,5)	(0,6)	(0,7)	(0,8)
(1,0)	(1,1)	(1,2)	(1,3)	(1,4)	(1,5)	(1,6)	(1,7)	(1,8)
(2,0)	(2,1)	(2,2)	(2,3)	(2,4)	(2,5)	(2,6)	(2,7)	(2,8)
(3,0)	(3,1)	(3,2)	(3,3)	(3,4)	(3,5)	(3,6)	(3,7)	(3,8)
(4,0)	(4,1)	(4,2)	(4,3)	(4,4)	(4,5)	(4,6)	(4,7)	(4,8)
(5,0)	(5,1)	(5,2)	(5,3)	(5,4)	(5,5)	(5,6)	(5,7)	(5,8)
(6,0)	(6,1)	(6,2)	(6,3)	(6,4)	(6,5)	(6,6)	(6,7)	(6,8)
(7,0)	(7,1)	(7,2)	(7,3)	(7,4)	(7,5)	(7,6)	(7,7)	(7,8)
(8,0)	(8,1)	(8,2)	(8,3)	(8,4)	(8,5)	(8,6)	(8,7)	(8,8)
=end
#*************
#縦チェック
#*************
    a = kekka.errchk(kekka.sdkarea,i,0,i,8)
#
#*************
#横チェック
#*************
    b = kekka.errchk(kekka.sdkarea,0,j,8,j)
#
#*************
#四角チェック
#*************
    if i <= 2
      if j <= 2
        c = kekka.errchk(kekka.sdkarea,0,0,2,2)
      elsif j <= 5
        c = kekka.errchk(kekka.sdkarea,0,3,2,5)
      else
        c = kekka.errchk(kekka.sdkarea,0,6,2,8)
      end
    elsif j <= 5
      if j <= 2
        c = kekka.errchk(kekka.sdkarea,3,0,5,2)
      elsif j <= 5
        c = kekka.errchk(kekka.sdkarea,3,3,5,5)
      else
        c = kekka.errchk(kekka.sdkarea,3,6,5,8)
      end
    else
      if j <= 2
        c = kekka.errchk(kekka.sdkarea,6,0,8,2)
      elsif j <= 5
        c = kekka.errchk(kekka.sdkarea,6,3,8,5)
      else
        c = kekka.errchk(kekka.sdkarea,6,6,8,8)
      end
    end
#
    print "a=",a," b=",b," c=",c,"\n"
    print " chk=",kekka.sdkarea,"\n"
#   print " j=",j,"\n"
#

    if a == ERR_OK and b == ERR_OK and c == ERR_OK
      j += 1
    end

    if j > 8
      j = 0
      break
    end
    
  end
 
  i += 1
  if i > 8
    break
  end

end
#
#ランダム
#
#縦チェック

