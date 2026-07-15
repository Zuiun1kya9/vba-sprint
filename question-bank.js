(function(){
'use strict';
const bank=[];
const add=(id,c,l,q,a,e,code)=>bank.push({id,c,l,q,a,ok:0,e,code:code||null});
const col=n=>{let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s};

// 1. 変数・演算子（50問）
for(let i=1;i<=50;i++){
  const a=2+i%11,b=1+(i*3)%9,mode=i%5;
  if(mode===0)add(1000+i,'変数・演算子','基礎','実行後の n の値は？',[String(a+b),String(a-b),String(a*b),String(b)],'加算代入の結果を求めます。',`Dim n As Long\nn = ${a}\nn = n + ${b}`);
  if(mode===1)add(1000+i,'変数・演算子','標準','整数除算「\\」の結果は？',[String(Math.floor(a*b/(b+1))),String(a*b/(b+1)),String((a*b)%(b+1)),'エラー'],`\\ は小数部分を除いた整数除算です。`, `Debug.Print ${a*b} \\ ${b+1}`);
  if(mode===2)add(1000+i,'変数・演算子','標準','Mod演算子の実行結果は？',[String((a+b)%b),String(Math.floor((a+b)/b)),String(a+b),String(b)],'Mod は割り算の余りを返します。',`Debug.Print ${a+b} Mod ${b}`);
  if(mode===3)add(1000+i,'変数・演算子','基礎','文字列を連結した結果は？',[`"VBA${a}"`,`"VBA ${a}"`,`"${a}VBA"`,'型不一致'],'& 演算子は文字列を連結します。',`Debug.Print "VBA" & ${a}`);
  if(mode===4)add(1000+i,'変数・演算子','標準','次の変数 x のデータ型は？',['Variant','Long','Integer','String'],'1つの Dim 文でも、As句を省略した変数は Variant です。',`Dim x, y As Long`);
}

// 2. 条件分岐・繰り返し（50問）
for(let i=1;i<=50;i++){
  const start=1+i%4,end=start+2+i%5,step=1+i%3,mode=i%5;
  if(mode===0){let n=0;for(let x=start;x<=end;x+=step)n++;add(1050+i,'条件分岐・繰り返し','標準','Debug.Printは何回実行される？',[`${n}回`,`${end-start+1}回`,`${step}回`,`${n+1}回`],'Forは開始値から終了値までStepずつ増加します。',`For i = ${start} To ${end} Step ${step}\n    Debug.Print i\nNext i`)}
  if(mode===1){const x=40+i;add(1050+i,'条件分岐・繰り返し','基礎','表示される文字列は？',[x>=60?'合格':'再試験',x>=60?'再試験':'合格','何も表示されない','エラー'],'Ifの条件がTrueならThen側、FalseならElse側を実行します。',`score = ${x}\nIf score >= 60 Then\n MsgBox "合格"\nElse\n MsgBox "再試験"\nEnd If`)}
  if(mode===2){const x=i%4;add(1050+i,'条件分岐・繰り返し','標準','Select Caseで表示される値は？',[x===0?'A':x<=2?'B':'C',x===0?'B':'A','C','何も表示されない'],'最初に一致したCaseだけが実行されます。',`n = ${x}\nSelect Case n\n Case 0: MsgBox "A"\n Case 1 To 2: MsgBox "B"\n Case Else: MsgBox "C"\nEnd Select`)}
  if(mode===3)add(1050+i,'条件分岐・繰り返し','基礎','条件を先に判定するループは？',['Do While 条件 ... Loop','Do ... Loop While 条件','For Eachだけ','With ... End With'],'Do Whileを先頭に置く形式は、処理前に条件を判定します。');
  if(mode===4)add(1050+i,'条件分岐・繰り返し','標準','Forループを途中で抜けるステートメントは？',['Exit For','End For','Break','Stop For'],'Exit Forは最も内側のForループを終了します。');
}

// 3. Range・Cells・Offset（50問）
for(let i=1;i<=50;i++){
  const r=1+i%20,c=1+(i*7)%12,dr=i%4,dc=i%3,mode=i%5,address=col(c)+r;
  if(mode===0)add(1100+i,'セル・Range','基礎',`Cells(${r}, ${c})が表すセルは？`,[address,col(r)+c,col(c+1)+r,col(c)+(r+1)],'Cellsは Cells(行番号, 列番号) の順です。');
  if(mode===1)add(1100+i,'セル・Range','標準','Offset後のセルは？',[col(c+dc)+(r+dr),col(c+dr)+(r+dc),col(c)+r,col(c+dc+1)+(r+dr+1)],'Offset(行方向, 列方向)で基準セルから移動します。',`Range("${address}").Offset(${dr}, ${dc}).Select`);
  if(mode===2)add(1100+i,'セル・Range','標準','Resizeで作られる範囲の大きさは？',[`${dr+2}行×${dc+2}列`,`${dc+2}行×${dr+2}列`,`${dr+1}行×${dc+1}列`,'1セル'],'Resize(行数, 列数)で範囲のサイズを指定します。',`Range("${address}").Resize(${dr+2}, ${dc+2}).Select`);
  if(mode===3)add(1100+i,'セル・Range','基礎','列全体を表す正しい記述は？',[`Columns("${col(c)}")`,`Rows("${col(c)}")`,`Column(${c})`,`Cells.Column(${c})`],'Columnsプロパティは指定した列全体を返します。');
  if(mode===4)add(1100+i,'セル・Range','標準','連続した表範囲を取得するプロパティは？',['CurrentRegion','UsedRangeOnly','EntireTable','DataArea'],'CurrentRegionは空白行・空白列で囲まれた連続範囲を返します。');
}

// 4. 配列・コレクション（50問）
for(let i=1;i<=50;i++){
  const lower=i%2,upper=lower+2+i%6,mode=i%5;
  if(mode===0)add(1150+i,'配列','基礎','配列の要素数はいくつ？',[String(upper-lower+1),String(upper),String(upper-lower),String(upper+1)],'下限と上限の両端を含むため、要素数は上限－下限＋1です。',`Dim arr(${lower} To ${upper}) As Long`);
  if(mode===1)add(1150+i,'配列','標準','配列の上限を返す関数は？',['UBound','LBound','UCase','Bound'],'UBoundは指定次元の添字の最大値を返します。');
  if(mode===2)add(1150+i,'配列','標準','既存の値を残して動的配列を拡張する記述は？',[`ReDim Preserve arr(${upper})`,`ReDim arr(${upper})`,`Resize Preserve arr(${upper})`,`Dim Preserve arr(${upper})`],'ReDim Preserveで既存要素を保持したまま最終次元を変更します。');
  if(mode===3)add(1150+i,'配列','基礎','配列の全要素を初期化するステートメントは？',['Erase','Clear','Reset Array','Delete'],'Eraseは固定配列を初期化し、動的配列の領域を解放します。');
  if(mode===4)add(1150+i,'配列','標準','コレクションの全要素を順番に処理するのに適した構文は？',['For Each ... Next','Select Case','Do Untilだけ','With ... End With'],'For Eachはコレクションや配列の各要素を列挙します。');
}

// 5. 文字列・日付・関数（50問）
for(let i=1;i<=50;i++){
  const word=['Excel','Standard','Macro','Range','Worksheet'][i%5],n=1+i%3,mode=i%5;
  if(mode===0)add(1200+i,'関数','基礎','Len関数の戻り値は？',[String(word.length),String(word.length-1),String(word.length+1),'0'],'Lenは文字列の文字数を返します。',`Debug.Print Len("${word}")`);
  if(mode===1)add(1200+i,'関数','基礎','Left関数の戻り値は？',[word.slice(0,n),word.slice(-n),word.slice(n),String(n)],'Left(文字列, 文字数)は左端から指定文字数を返します。',`Debug.Print Left("${word}", ${n})`);
  if(mode===2)add(1200+i,'関数','基礎','Right関数の戻り値は？',[word.slice(-n),word.slice(0,n),word.slice(n),String(n)],'Right(文字列, 文字数)は右端から指定文字数を返します。',`Debug.Print Right("${word}", ${n})`);
  if(mode===3)add(1200+i,'関数','標準','文字列内の位置を検索する関数は？',['InStr','FindText','Search','IndexOf'],'InStrは文字列中で別の文字列が最初に現れる位置を返します。');
  if(mode===4)add(1200+i,'関数','標準','現在の日付だけを返すVBA関数は？',['Date','Now','Time','Today'],'Dateは現在の日付、Nowは現在の日付と時刻を返します。');
}

// 6. プロシージャ・引数（50問）
for(let i=1;i<=50;i++){
  const n=2+i%8,mode=i%5;
  if(mode===0)add(1250+i,'プロシージャ','標準','実行後の n の値は？',[String(n+1),String(n),String(n*2),'エラー'],'ByRefは呼び出し元の変数自体を参照するため、変更が反映されます。',`Sub Main()\n Dim n As Long: n = ${n}\n AddOne n\n Debug.Print n\nEnd Sub\nSub AddOne(ByRef x As Long)\n x = x + 1\nEnd Sub`);
  if(mode===1)add(1250+i,'プロシージャ','標準','実行後の n の値は？',[String(n),String(n+1),String(n*2),'エラー'],'ByValは値のコピーを渡すため、呼び出し元の変数は変わりません。',`Dim n As Long: n = ${n}\nCall AddOne(n)\nSub AddOne(ByVal x As Long)\n x = x + 1\nEnd Sub`);
  if(mode===2)add(1250+i,'プロシージャ','基礎','戻り値を返せるプロシージャは？',['Function','Sub','Module','Call'],'Functionは自身の名前へ代入した値を戻り値として返します。');
  if(mode===3)add(1250+i,'プロシージャ','標準','Callを使って引数を渡す正しい記述は？',[`Call Test(${n})`,`Call Test ${n}`,`Test(${n})`,`Run Test(${n})`],'Callを使う場合、引数リストを丸かっこで囲みます。');
  if(mode===4)add(1250+i,'プロシージャ','標準','プロシージャを同一モジュール内だけから呼べる宣言は？',['Private Sub','Public Sub','Global Sub','Static Module'],'Privateで宣言したプロシージャは、そのモジュール内だけで利用できます。');
}

// 7. ブック・シート・イベント（50問）
for(let i=1;i<=50;i++){
  const sheet=1+i%5,mode=i%5;
  if(mode===0)add(1300+i,'ブック・シート','基礎','コードが保存されているブックを表すのは？',['ThisWorkbook','ActiveWorkbook','Workbooks(1)','CurrentWorkbook'],'ThisWorkbookは実行中のコードを含むブックです。');
  if(mode===1)add(1300+i,'ブック・シート','標準',`左から${sheet}番目のワークシートを参照する記述は？`,[`Worksheets(${sheet})`,`Worksheet(${sheet})`,`Sheets.Item(0)`,`Worksheets[${sheet}]`],'Worksheetsコレクションは1から始まるインデックスで参照できます。');
  if(mode===2)add(1300+i,'ブック・シート','標準','ブックを保存せずに閉じる記述は？',['ActiveWorkbook.Close SaveChanges:=False','ActiveWorkbook.Close Save:=No','Workbooks.Close Falseだけ','Application.Quit False'],'CloseメソッドのSaveChanges引数へFalseを指定します。');
  if(mode===3)add(1300+i,'イベント','標準','セルの値が変更されたときのシートイベントは？',['Worksheet_Change','Worksheet_Select','Workbook_Change','Cell_Change'],'Worksheet_Changeはワークシート上のセル変更時に発生します。');
  if(mode===4)add(1300+i,'イベント','標準','ブックを開いたときに発生するイベントは？',['Workbook_Open','Workbook_Start','Application_Open','Book_ActivateOnly'],'Workbook_Openはブックを開いたときに発生します。');
}

// 8. 検索・フィルター・並べ替え・テーブル（50問）
for(let i=1;i<=50;i++){
  const field=1+i%5,mode=i%5;
  if(mode===0)add(1350+i,'データ操作','標準','Range.Findが見つからなかった場合の戻り値は？',['Nothing','False','0','空文字'],'Findメソッドは該当セルがない場合にNothingを返します。');
  if(mode===1)add(1350+i,'データ操作','標準',`${field}列目を「東京」で絞り込む記述は？`,[`Range("A1").CurrentRegion.AutoFilter Field:=${field}, Criteria1:="東京"`,`AutoFilter Column:=${field}, Value:="東京"`,`Range("A1").Filter ${field}, "東京"`,`Columns(${field}).Hidden = "東京"`],'AutoFilterのFieldは対象範囲内での列位置です。');
  if(mode===2)add(1350+i,'データ操作','標準','昇順を表すExcel定数は？',['xlAscending','xlDescending','xlUp','xlTopToBottom'],'xlAscendingは昇順、xlDescendingは降順を表します。');
  if(mode===3)add(1350+i,'データ操作','標準','ワークシート上のテーブルを表すオブジェクトは？',['ListObject','TableObject','DataTable','ListRange'],'ExcelのテーブルはListObjectオブジェクトで表されます。');
  if(mode===4)add(1350+i,'データ操作','標準','テーブルのデータ部分だけを表すプロパティは？',['DataBodyRange','CurrentRegion','TableData','BodyCells'],'DataBodyRangeは見出しや集計行を除くデータ部分を返します。');
}

// 9. エラー処理・デバッグ（50問）
for(let i=1;i<=50;i++){
  const mode=i%5;
  if(mode===0)add(1400+i,'エラー・デバッグ','基礎','実行時エラー発生時に指定ラベルへ移動する構文は？',['On Error GoTo ラベル','Try Catch','On Error Resume ラベル','GoTo Error'],'On Error GoToでエラー処理ラベルを指定します。');
  if(mode===1)add(1400+i,'エラー・デバッグ','標準','現在のエラー情報を保持するオブジェクトは？',['Err','ErrorInfo','Exception','Debug'],'ErrオブジェクトのNumberやDescriptionでエラー情報を取得できます。');
  if(mode===2)add(1400+i,'エラー・デバッグ','標準','設定済みのエラーハンドラーを無効にする記述は？',['On Error GoTo 0','On Error Off','Err.Clearだけ','Resume 0'],'On Error GoTo 0は現在のプロシージャのエラーハンドラーを無効にします。');
  if(mode===3)add(1400+i,'エラー・デバッグ','基礎','イミディエイトウィンドウに値を出力する命令は？',['Debug.Print','Console.Write','Print.Debug','MsgBox.Print'],'Debug.Printは式の値をイミディエイトウィンドウへ出力します。');
  if(mode===4)add(1400+i,'エラー・デバッグ','基礎','コードを1行ずつ実行するデバッグ方法は？',['ステップ実行','コンパイル','ウォッチ式だけ','マクロ記録'],'ステップ実行では処理を1行ずつ進めて状態を確認できます。');
}

// 10. ファイル・Application・その他（50問）
for(let i=1;i<=50;i++){
  const mode=i%5;
  if(mode===0)add(1450+i,'ファイル・その他','標準','ファイルやフォルダーの存在確認に使える関数は？',['Dir','Exists','FileFind','PathCheck'],'Dir関数は条件に一致するファイル名やフォルダー名を返します。');
  if(mode===1)add(1450+i,'ファイル・その他','標準','テキストファイルを追記モードで開く指定は？',['Open パス For Append As #番号','Open パス For Output As #番号','Open パス For Input As #番号','File.Open Append'],'Appendは既存内容の末尾へ追記するモードです。');
  if(mode===2)add(1450+i,'ファイル・その他','標準','画面更新を停止する記述は？',['Application.ScreenUpdating = False','Application.Visible = False','Screen.Refresh = False','Workbook.Update = False'],'ScreenUpdatingをFalseにするとマクロ実行中の画面更新を抑制できます。');
  if(mode===3)add(1450+i,'ファイル・その他','標準','警告ダイアログを一時的に非表示にする記述は？',['Application.DisplayAlerts = False','Application.Alerts = None','MsgBox.Enabled = False','Warnings.Hide'],'DisplayAlertsをFalseにすると特定の確認ダイアログを抑制できます。終了前にTrueへ戻します。');
  if(mode===4)add(1450+i,'ファイル・その他','標準','Withブロックを使う主な目的は？',['同じオブジェクトへの記述をまとめる','変数を宣言する','エラーを無視する','ループを終了する'],'Withを使うと同じオブジェクト名の繰り返しを省略できます。');
}

window.EXTRA_QUESTIONS=bank;
})();
