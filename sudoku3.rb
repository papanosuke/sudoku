# -*- encoding: utf-8 -*-
#*************************************************************
#*数独問題作成
#*************************************************************
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
①ランダムで1～9の数字をchoiceareaから選択する
②縦、横、四角、続行チェックの４通りのチェックを実施する
③エラーがなかったらchoiceareaから削除
④①～③の繰り返し

=end

ERR_OK = 0
ERR_NG = 1
HANTEI1 = 5
HANTEI2 = 8

#*************************************************************
#*クラス
#*************************************************************
class Array
#配列からランダムで選択するインスタンスを作成
  def choice
    at( rand( size ) )
  end
end

class Sudoku
  attr_accessor :sdkarea
  attr_accessor :sdkarea2
  attr_accessor :sdkarea3
  attr_accessor :choicearea
  attr_accessor :levelnum
  attr_accessor :st_x
  attr_accessor :ed_x
  attr_accessor :st_y
  attr_accessor :ed_y
  attr_accessor :saikeisan_flg
  
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
  def initialize()
    @sdkarea = self.secondarray(9,9)
    for i in 0..8 do
      for j in 0..8 do
        @sdkarea[i][j] = 0
      end
    end

    @choicearea = self.secondarray(9,9)
    for i in 0..8 do
      for j in 0..8 do
        @choicearea[i][j] = j + 1
      end
####p @choicearea[i]
    end

####p @choicearea[i]

    @st_x = 0
    @st_y = 0
    @ed_x = 0
    @ed_y = 0

    @levelnum = 0
    @saikeisan_flg = 0

  end
#********************
#*sdkarea再初期化
#********************
  def sdkareaclear(sarea,gyo)
    for j in 0..8 do
      sarea[gyo][j] = 0
    end
#確認
print "*-------- sdkareaクリア後 --------*\n"
   for j in 0..8
     print @sdkarea[j],"\n"
   end

  end
#
#********************
#*choicearea再初期化
#********************
  def choiceareaclear(carea,gyo)
    for j in 0..8 do
      carea[gyo][j] = j + 1
    end

#確認
print "*-------- choiceareaクリア後 --------*\n"
   for j in 0..8
     print @choicearea[j],"\n"
   end

  end
#
#********************
#エラーチェック
#①重複チェック
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
#エラーあり＝ERR_NGを返す
        return ERR_NG if chkarea[sdkarea[i][j] - 1] > 1
      end

    end

#エラーなし＝ERR_OKを返す
    return ERR_OK
  end

#********************
#エラーチェック
#②続行チェック
#********************
#(1,0)(2,0)(3,0)の数字が(5,0)(5,1)(5,2)の残りの数字だと処理がループになる
  def errchk2(sdkarea, st_x, st_y)


    carea = [1,2,3,4,5,6,7,8,9]
    for i in st_x-2..(st_x-1)
      for j in st_y..(st_y+2)
#       print "a=",sdkarea[i][j]," i=",i," j=",j,"\n"
        carea.delete_if{|x| x == sdkarea[i][j]}
#       p carea
      end
    end

    if st_x == 5

      for j in st_y..(st_y+2)

        chk = []

        for i in (st_x - HANTEI1)..(st_x - HANTEI1 + 2)
          chk << sdkarea[i][j]
        end
        chk.sort!

########print "carea=",carea," chk=",chk,"\n"

        if carea == chk
          return ERR_NG
        end

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

print "*----- 問題作成開始 -----*\n"
print "開始時刻=",Time.now,"\n"

kekka = Sudoku::new
#********************
#レベル設定
#********************
while true
  print "*----------------------------------------------------------------------------*\n"
  print " レベルは？[1-4]\n"
  print "*----------------------------------------------------------------------------*\n"
  print "==> "
  kekka.levelnum = gets.to_i
  if kekka.levelnum >= 1 and kekka.levelnum <= 4
    break
  end
end
#

#********************
#数字埋め作業
#********************
i = 0

try = 1
#
while true
#
  j = 0

  while true
#
#数字をランダムで生成(choiceareaから選択)
    chkflg = ERR_OK
#
#縦で使用されているものは削除しておく
    if i > 0
      for h in 0..i-1 do
        kekka.choicearea[j].delete_if {|x| x == kekka.sdkarea[h][j]}
      end
    end
#
    if kekka.choicearea[j].length == 0
      print "choicearea要素数=0"," gyo=",j,"\n"
      break
    end
#
    kekka.sdkarea[i][j] = kekka.choicearea[j].choice
#
    j += 1
#
    if j > 8
      break
    end
#
  end
#
#
#===>DEBUG START
# if kekka.saikeisan_flg == 1
#   print "i=",i," j=",j," suji=",kekka.sdkarea[i],"\n"
# end
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
#********************
#横チェック
#********************
  a = kekka.errchk(kekka.sdkarea,i,0,i,8)
#
#********************
#縦チェック
#********************
  b = []
  for k in 0..8
    b[k] = kekka.errchk(kekka.sdkarea,0,k,8,k)
  end
#
#********************
#四角チェック
#********************
  c = []
  if i <= 2
    c[0] = kekka.errchk(kekka.sdkarea,0,0,2,2)
    c[1] = kekka.errchk(kekka.sdkarea,0,3,2,5)
    c[2] = kekka.errchk(kekka.sdkarea,0,6,2,8)
  elsif i <= 5
    c[0] = kekka.errchk(kekka.sdkarea,3,0,5,2)
    c[1] = kekka.errchk(kekka.sdkarea,3,3,5,5)
    c[2] = kekka.errchk(kekka.sdkarea,3,6,5,8)
  else
    c[0] = kekka.errchk(kekka.sdkarea,6,0,8,2)
    c[1] = kekka.errchk(kekka.sdkarea,6,3,8,5)
    c[2] = kekka.errchk(kekka.sdkarea,6,6,8,8)
  end
#
#********************
#続行チェック
#********************
#(1,0)(2,0)(3,0)の数字が(5,0)(5,1)(5,2)の残りの数字だと処理がループになる
  d = []
  if i == 5
    d[0] = kekka.errchk2(kekka.sdkarea,5,0)
    d[1] = kekka.errchk2(kekka.sdkarea,5,3)
    d[2] = kekka.errchk2(kekka.sdkarea,5,6)
  end
#
# print "a=",a," b=",b," c=",c,"\n"
# print " kekka=",kekka.sdkarea,"\n"
#
#
#********************
#すべてのエラー確認
#********************
#
  chkflg = ERR_OK
  if a != ERR_OK
      chkflg = ERR_NG
  end

  for m in 0..8
    if b[m] != ERR_OK
      chkflg = ERR_NG
    end
  end

  for m in 0..2
    if c[m] != ERR_OK
      chkflg = ERR_NG
    end
  end

  err_z = ERR_OK
  if i == 5
    for m in 0..2
      if d[m] != ERR_OK
        err_z = ERR_NG
      end
    end
  end
  if err_z == ERR_NG
     kekka.choiceareaclear(kekka.choicearea,i)
     kekka.choiceareaclear(kekka.choicearea,i-1)
     kekka.choiceareaclear(kekka.choicearea,i-2)
     kekka.choiceareaclear(kekka.choicearea,i-3)
     kekka.choiceareaclear(kekka.choicearea,i-4)
     kekka.choiceareaclear(kekka.choicearea,i-5)
     kekka.sdkareaclear(kekka.sdkarea,i)
     kekka.sdkareaclear(kekka.sdkarea,i-1)
     kekka.sdkareaclear(kekka.sdkarea,i-2)
     kekka.sdkareaclear(kekka.sdkarea,i-3)
     kekka.sdkareaclear(kekka.sdkarea,i-4)
     kekka.sdkareaclear(kekka.sdkarea,i-5)
     print "iをマイナスします！",i,"⇒",i-5,"\n"
     i -= 5
     kekka.saikeisan_flg = 1
  end
#
#
#********************
#choiceareaから削除
#********************
  if chkflg == ERR_OK
#===>debug start
    print "ERR_OK...",i,"\n"
#<===debug end
    for h in 0..8 do
#===>debug start
#####p kekka.sdkarea[i][h]
#<===debug end
      kekka.choicearea[h].delete_if {|x| x == kekka.sdkarea[i][h]}
#===>debug start
#####p kekka.choicearea[h]
#<===debug end
    end
#===>debug start
#   print kekka.sdkarea[i],"\n"
#<===debug end
    i += 1

    try = 0
  else
    try += 1

    for h in 0..8 do
      kekka.sdkarea[i][h] = 0
    end
  end

  try_disp = try.to_s.reverse
  if try_disp[0..4] == "00000"
    print "try回数=",try,"\n"
#   for h in 0..8
#     print kekka.sdkarea[h],"\n"
#   end
  end

##if try > 100000
##  print "問題の作成に失敗しました。","\n"
##  exit!
##end

  if i > 8
    break
  end

end
#
#********************
#確認表示
#********************
print "*-------- 確認表示① --------*\n"
for h in 0..8
  print kekka.sdkarea[h],"\n"
end
#
#********************
#ランダム空白作成
#********************
kuhaku_max = 0
case kekka.levelnum
  when 1
   kuhaku_max = 35
  when 2
   kuhaku_max = 40
  when 3
   kuhaku_max = 45
  when 4
   kuhaku_max = 50
end

##kuhaku = 0
##while true
#   3.step(9,3) { |i|
#     3.step(9,3) { |j|
#       rnd_i = rand(i)
#       rnd_j = rand(j)
#       kuhaku_num = kekka.sdkarea[rnd_i][rnd_j]
#       if kuhaku_num != 0
#         kekka.sdkarea[rnd_i][rnd_j] = 0
#         kuhaku += 1
#       end
#     }
#   }
#   break if kuhaku >= kuhaku_max
# end
kuhaku = 0
while
  i = rand(9)
  j = rand(9)
  if kekka.sdkarea[i][j] != 0
    kekka.sdkarea[i][j] = 0
    kuhaku += 1
  end
  break if kuhaku == kuhaku_max
end
#
#********************
#回答保存
#********************
#deep copy
kekka.sdkarea2 = Marshal.load(Marshal.dump(kekka.sdkarea))
#********************
#確認表示
#********************
print "*-------- 確認表示② --------*\n"
for h in 0..8
  print kekka.sdkarea2[h],"\n"
end
#
#********************
#回答確認
#********************
#①昇順バックトラック
kuhaku = kuhaku_max
mae_kuhaku = -1
while kuhaku != 0
  i = 0
  while true
#
    j = 0
#
    while true
#
      if kekka.sdkarea2[i][j] == 0
#
        kaito = 0
        kaito_bk = 0
#
#順番に１から９まで当てはめていく
        for m in 1..9
#
          kekka.sdkarea2[i][j] = m
          a = kekka.errchk(kekka.sdkarea2,i,0,i,8)
          b = kekka.errchk(kekka.sdkarea2,0,j,8,j)
#
          if i <= 2
            if i <= 2
              c = kekka.errchk(kekka.sdkarea2,0,0,2,2)
            elsif i <= 5
              c = kekka.errchk(kekka.sdkarea2,0,3,2,5)
            else
              c = kekka.errchk(kekka.sdkarea2,0,6,2,8)
            end
          elsif i <= 5
            if i <= 2
              c = kekka.errchk(kekka.sdkarea2,3,0,5,2)
            elsif i <= 5
              c = kekka.errchk(kekka.sdkarea2,3,3,5,5)
            else
              c = kekka.errchk(kekka.sdkarea2,3,6,5,8)
            end
          else
            if i <= 2
              c = kekka.errchk(kekka.sdkarea2,6,0,8,2)
            elsif i <= 5
              c = kekka.errchk(kekka.sdkarea2,6,3,8,5)
            else
              c = kekka.errchk(kekka.sdkarea2,6,6,8,8)
            end
          end
#
          if a == ERR_OK and b == ERR_OK and c == ERR_OK
            kaito += 1
            kaito_bk = m
          end
#
        end
#
#複数回答があるものはゼロに戻す。１つのものは解答と断定
        if kaito > 1
          kekka.sdkarea2[i][j] = 0
        elsif kaito == 0
          kekka.sdkarea2[i][j] = 0
        else
          kuhaku -= 1
          kekka.sdkarea2[i][j] = kaito_bk
        end
#
      end
#
      j += 1
      if j > 8
        break
      end
#
    end
#
    i += 1
#
    if i > 8
      break
    end
#
  end
#
  if kuhaku == mae_kuhaku
    print "空白=",kuhaku,"\n"
    print "空白が減らないため処理を中断します！\n"
    exit!
  end
  print "空白=",kuhaku,"\n"
  mae_kuhaku = kuhaku
#
#
end

#********************
#ゼロの個数を確認
#********************
##kuhaku = 0
##for i in 0..8
##  for j in 0..8
##    if  kekka.sdkarea2[i][j] == 0
##      kuhaku +=1
##    end
##  end
##end
##print "ゼロの個数=",kuhaku,"個","\n"
##
#********************
#確認表示
#********************
print "*-------- 確認表示④ --------*\n"
for h in 0..8
  print kekka.sdkarea[h],"\n"
end
print "*-------- 確認表示⑤ --------*\n"
for h in 0..8
  print kekka.sdkarea2[h],"\n"
end
#
#********************
#エクセル書き出し
#********************
openExcelWorkbook('main.xlsx') do |book|
  i = 1
  mondai = book.Worksheets.Item('sudoku_m')
  kaito = book.Worksheets.Item('sudoku_k')

#オブジェクトのメソッド表示
##p sheet.ole_methods

  for i in 0..8
    for j in 0..8
      if kekka.sdkarea[i][j] == 0
        mondai.Cells(i+2,j+1).Value = ""
        mondai.Cells(i+2,j+1).Font.ColorIndex = 3
      else
        mondai.Cells(i+2,j+1).Value = kekka.sdkarea[i][j]
        mondai.Cells(i+2,j+1).Font.ColorIndex = 1
      end
    end
  end

  for i in 0..8
    for j in 0..8
      kaito.Cells(i+2,j+1).Value = kekka.sdkarea2[i][j]
      kaito.Cells(i+2,j+1).Font.ColorIndex = 1
    end
  end

  book.save
end
#********************
#終了表示
#********************
print "終了時刻=",Time.now,"\n"
