(function(){
'use strict';
const bank=[];
const add=(id,c,l,q,a,e,code)=>{
  if(new Set(a).size!==a.length)console.error(`Duplicate choices: ${id}`);
  bank.push({id,c,l,q,a,ok:0,e,code:code||null});
};

// 1. Accessオブジェクトモデル（50問）
for(let i=1;i<=50;i++){
  const mode=i%5,form=`F_受注_${i}`,report=`R_売上_${i}`;
  if(mode===0)add(2000+i,'オブジェクトモデル','基礎',`変数 db${i} に現在のAccessデータベースを代入する式は？`,['CurrentDb','CurrentProject','ActiveDatabase','Application.Database'],'CurrentDb関数は、現在開いているデータベースを表すDAO.Databaseオブジェクトを返します。',`Dim db${i} As DAO.Database\nSet db${i} = CurrentDb`);
  if(mode===1)add(2000+i,'オブジェクトモデル','標準','現在のAccessファイルの完全なパスを取得するプロパティは？',['CurrentProject.FullName','CurrentDb.FullName','Application.PathName','CurrentProject.Directory'],'CurrentProject.FullNameは、現在のプロジェクトのパスとファイル名を返します。',`Debug.Print CurrentProject.FullName  '確認 ${i}`);
  if(mode===2)add(2000+i,'オブジェクトモデル','標準','AccessObjectとしてテーブル一覧を参照するコレクションは？',['CurrentData.AllTables','CurrentProject.AllTables','CurrentDb.Tables','Application.TableObjects'],'CurrentData.AllTablesは、現在のデータベースにあるテーブルをAccessObjectとして列挙します。',`Debug.Print CurrentData.AllTables(${i%3}).Name`);
  if(mode===3)add(2000+i,'オブジェクトモデル','基礎',`フォーム「${form}」が開いているか調べる式は？`,[`CurrentProject.AllForms("${form}").IsLoaded`,`Forms("${form}").IsOpen`,`CurrentData.Forms("${form}").Loaded`,`DoCmd.IsOpen("${form}")`],'CurrentProject.AllForms内のAccessObjectが持つIsLoadedプロパティで、フォームが開いているか確認できます。');
  if(mode===4)add(2000+i,'オブジェクトモデル','標準',`レポート「${report}」をデザインビューで選択するDoCmdメソッドは？`,[`DoCmd.SelectObject acReport, "${report}", True`,`DoCmd.OpenObject acReport, "${report}"`,`Application.Select acReport, "${report}"`,`Reports.Select "${report}"`],'SelectObjectメソッドは、オブジェクトの種類と名前を指定してデータベースウィンドウ上のオブジェクトを選択します。');
}

// 2. テーブル・フィールド・インデックス（50問）
for(let i=1;i<=50;i++){
  const mode=i%5,table=`T_顧客_${i}`,field=`備考${i}`,index=`PK_${i}`;
  if(mode===0)add(2050+i,'テーブル・フィールド','基礎',`テーブル「${table}」のTableDefを取得する記述は？`,[`CurrentDb.TableDefs("${table}")`,`CurrentDb.Tables("${table}")`,`CurrentProject.TableDefs("${table}")`,`DoCmd.Table("${table}")`],'DAO.DatabaseのTableDefsコレクションをテーブル名で参照するとTableDefを取得できます。');
  if(mode===1)add(2050+i,'テーブル・フィールド','標準',`DAOで短いテキスト型のフィールド「${field}」を作るデータ型定数は？`,['dbText','dbString','acText','dbMemo'],'DAOの短いテキスト型はdbTextです。長いテキスト型にはdbMemoを使用します。',`Set fld = tdf.CreateField("${field}", dbText, ${20+i})`);
  if(mode===2)add(2050+i,'テーブル・フィールド','標準','新しいFieldをTableDefへ追加するメソッドは？',['tdf.Fields.Append fld','tdf.Fields.Add fld','tdf.AppendField fld','db.Fields.Insert fld'],'作成したFieldはTableDefのFieldsコレクションへAppendします。',`Set fld = tdf.CreateField("${field}", dbText, ${10+i})\ntdf.Fields.Append fld`);
  if(mode===3)add(2050+i,'テーブル・フィールド','標準',`インデックス「${index}」を主キーに設定するプロパティは？`,['idx.Primary = True','idx.Unique = Primary','idx.Key = True','idx.Attributes = dbPrimary'],'DAO.IndexのPrimaryプロパティをTrueにすると、そのインデックスが主キーになります。',`Set idx = tdf.CreateIndex("${index}")\nidx.Primary = True`);
  if(mode===4)add(2050+i,'テーブル・フィールド','基礎','TableDefsコレクションの内容を最新状態にするメソッドは？',['CurrentDb.TableDefs.Refresh','CurrentDb.TableDefs.Update','Application.RefreshTables','DoCmd.RefreshSchema'],'RefreshメソッドはTableDefsコレクションを最新の状態へ更新します。',`CurrentDb.TableDefs.Refresh  '更新 ${i}`);
}

// 3. クエリ・QueryDef（50問）
for(let i=1;i<=50;i++){
  const mode=i%5,query=`Q_抽出_${i}`,param=`pID${i}`;
  if(mode===0)add(2100+i,'クエリ・QueryDef','基礎',`保存クエリ「${query}」のQueryDefを取得する記述は？`,[`CurrentDb.QueryDefs("${query}")`,`CurrentDb.Queries("${query}")`,`CurrentProject.QueryDefs("${query}")`,`DoCmd.Query("${query}")`],'DAO.DatabaseのQueryDefsコレクションを名前で参照すると、保存済みQueryDefを取得できます。');
  if(mode===1)add(2100+i,'クエリ・QueryDef','標準','保存しない一時QueryDefを作成するときの名前は？',['空文字列（""）','Null','"Temporary"','省略できない'],'CreateQueryDefのName引数に空文字列を指定すると、QueryDefsコレクションへ保存されない一時QueryDefになります。',`Set qdf = CurrentDb.CreateQueryDef("", "SELECT * FROM T_${i}")`);
  if(mode===2)add(2100+i,'クエリ・QueryDef','標準',`パラメーター「${param}」へ値を代入する記述は？`,[`qdf.Parameters("${param}") = ${i}`,`qdf.Parameter("${param}").Value(${i})`,`qdf.${param} = ${i}`,`Parameters!${param}.Set ${i}`],'QueryDefのParametersコレクションを名前で参照し、値を代入します。');
  if(mode===3)add(2100+i,'クエリ・QueryDef','基礎','QueryDefに設定されたSQL文を取得・変更するプロパティは？',['SQL','Source','CommandText','QueryText'],'QueryDefのSQLプロパティは、そのクエリが実行するSQL文字列を保持します。',`Debug.Print qdf.SQL  'クエリ ${i}`);
  if(mode===4)add(2100+i,'クエリ・QueryDef','標準',`新しい保存クエリ「${query}」を作成する記述は？`,[`Set qdf = db.CreateQueryDef("${query}", sql)`,`Set qdf = db.CreateQueryDef("", sql)`,`db.QueryDefs.Add "${query}", sql`,`Set qdf = New DAO.QueryDef("${query}")`],'CreateQueryDefへ空でない名前を指定すると、作成したQueryDefがQueryDefsコレクションへ自動的に追加され保存されます。');
}

// 4. DAO Recordset（50問）
for(let i=1;i<=50;i++){
  const mode=i%5,table=`T_商品_${i}`,value=100+i;
  if(mode===0)add(2150+i,'DAO Recordset','基礎',`テーブル「${table}」を更新可能なダイナセットで開く定数は？`,['dbOpenDynaset','dbOpenSnapshot','dbOpenTableReadOnly','acOpenDynaset'],'dbOpenDynasetを指定すると、通常は編集可能な動的Recordsetを開きます。',`Set rs = CurrentDb.OpenRecordset("${table}", dbOpenDynaset)`);
  if(mode===1)add(2150+i,'DAO Recordset','標準','レコードが1件もないRecordsetでTrueになる組み合わせは？',['BOF = True かつ EOF = True','BOF = False かつ EOF = True','BOF = True かつ EOF = False','RecordCount = -1だけ'],'空のRecordsetではカレントレコードが存在せず、BOFとEOFがどちらもTrueになります。',`If rs.BOF And rs.EOF Then Debug.Print "空です ${i}"`);
  if(mode===2)add(2150+i,'DAO Recordset','基礎','Recordsetを先頭レコードへ移動するメソッドは？',['MoveFirst','MoveStart','First','GoTop'],'MoveFirstメソッドでRecordsetの先頭レコードをカレントレコードにします。',`rs.MoveFirst  '移動 ${i}`);
  if(mode===3)add(2150+i,'DAO Recordset','標準',`新しいレコードに値 ${value} を追加するときに必要な順序は？`,['AddNew → 値の代入 → Update','Edit → 値の代入 → Save','Insert → Update → 値の代入','New → Commit → 値の代入'],'DAO Recordsetへの新規追加はAddNewで編集バッファを作り、値を設定してUpdateで確定します。',`rs.AddNew\nrs!金額 = ${value}\nrs.Update`);
  if(mode===4)add(2150+i,'DAO Recordset','標準',`既存レコードの金額を ${value} に変更するときに必要な順序は？`,['Edit → 値の代入 → Update','AddNew → 値の代入 → Update','Update → 値の代入 → Edit','Modify → Save → Commit'],'既存レコードの編集はEditを呼び出し、フィールド値を変更してUpdateで確定します。',`rs.Edit\nrs!金額 = ${value}\nrs.Update`);
}

// 5. フォーム・コントロール・イベント（50問）
for(let i=1;i<=50;i++){
  const mode=i%5,form=`F_顧客_${i}`,control=`txt氏名${i}`;
  if(mode===0)add(2200+i,'フォーム・イベント','基礎','フォーム自身を参照するキーワードは？',['Me','ThisForm','CurrentForm','Self'],'フォームモジュール内のMeは、そのコードを含むフォームのインスタンスを参照します。',`Me.Caption = "顧客 ${i}"`);
  if(mode===1)add(2200+i,'フォーム・イベント','基礎',`開いているフォーム「${form}」のコントロール「${control}」を参照する記述は？`,[`Forms("${form}").Controls("${control}")`,`Form("${form}").Control("${control}")`,`CurrentProject.Forms("${form}").${control}`,`DoCmd.Controls("${control}")`],'Formsコレクションから開いているフォームを取得し、Controlsコレクションを名前で参照します。');
  if(mode===2)add(2200+i,'フォーム・イベント','標準','フォームの更新をBeforeUpdateイベントで取り消す記述は？',['Cancel = True','Me.Undo = False','Event.Cancel','Exit Sub = True'],'BeforeUpdateイベントのCancel引数へTrueを代入すると更新を取り消せます。',`Private Sub Form_BeforeUpdate(Cancel As Integer)\n    If IsNull(Me!${control}) Then Cancel = True\nEnd Sub`);
  if(mode===3)add(2200+i,'フォーム・イベント','標準','DoCmd.OpenFormで渡したOpenArgsをフォーム側で受け取るプロパティは？',['Me.OpenArgs','Me.Arguments','Form.Params','DoCmd.OpenArgs'],'OpenFormのOpenArgs引数は、開いたフォームのOpenArgsプロパティから取得します。',`Debug.Print Me.OpenArgs  'フォーム ${i}`);
  if(mode===4)add(2200+i,'フォーム・イベント','基礎',`テキストボックス「${control}」の値が更新された直後に発生するイベントは？`,['AfterUpdate','AfterChange','ValueChanged','OnCommit'],'コントロールのAfterUpdateイベントは、変更された値が更新された後に発生します。',`Private Sub ${control}_AfterUpdate()\n    Debug.Print Me!${control}\nEnd Sub`);
}

// 6. レポート（50問）
for(let i=1;i<=50;i++){
  const mode=i%5,report=`R_請求_${i}`,id=1000+i;
  if(mode===0)add(2250+i,'レポート','基礎',`レポート「${report}」を印刷プレビューで開く定数は？`,['acViewPreview','acViewReport','acViewDesign','acViewNormal'],'DoCmd.OpenReportのView引数にacViewPreviewを指定すると印刷プレビューで開きます。',`DoCmd.OpenReport "${report}", acViewPreview`);
  if(mode===1)add(2250+i,'レポート','標準',`ID=${id}のレコードだけに絞ってレポート「${report}」を開く引数は？`,['WhereCondition','FilterName','OpenArgs','DataMode'],'OpenReportのWhereCondition引数へWHERE句を除いた条件式を渡すと、表示レコードを絞り込めます。',`DoCmd.OpenReport "${report}", acViewPreview, , "ID=${id}"`);
  if(mode===2)add(2250+i,'レポート','標準','レポートにデータがないときに発生し、出力を中止できるイベントは？',['NoData','Empty','BeforePrint','NotFound'],'Report_NoDataイベントでCancelをTrueにすると、データがないレポートを開く処理を中止できます。',`Private Sub Report_NoData(Cancel As Integer)\n    Cancel = True  '確認 ${i}\nEnd Sub`);
  if(mode===3)add(2250+i,'レポート','基礎','レポートの現在のページ番号を表すプロパティは？',['Me.Page','Me.Pages','Me.PageNumber','Report.CurrentPage'],'Pageプロパティは現在のページ番号を返します。Pagesは総ページ数です。',`Me!txtPage = Me.Page  '表示 ${i}`);
  if(mode===4)add(2250+i,'レポート','標準','レポートの詳細セクションが書式設定されるときのイベントは？',['Detail_Format','Detail_Printing','Report_Render','Section_Style'],'Detail_Formatイベントは、詳細セクションが出力用に書式設定されるときに発生します。',`Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)\n    Me!txtNo = ${i}\nEnd Sub`);
}

// 7. DoCmd・入出力（50問）
for(let i=1;i<=50;i++){
  const mode=i%5,form=`F_検索_${i}`,table=`T_取込_${i}`;
  if(mode===0)add(2300+i,'DoCmd・入出力','基礎',`フォーム「${form}」を開くメソッドは？`,[`DoCmd.OpenForm "${form}"`,`DoCmd.ShowForm "${form}"`,`Application.Open "${form}"`,`Forms.Open "${form}"`],'DoCmd.OpenFormメソッドで名前を指定してフォームを開きます。');
  if(mode===1)add(2300+i,'DoCmd・入出力','標準',`フォーム「${form}」を保存して閉じる記述は？`,[`DoCmd.Close acForm, "${form}", acSaveYes`,`DoCmd.CloseForm "${form}", True`,`Forms("${form}").Close Save:=True`,`Application.Close acForm, "${form}"`],'DoCmd.Closeへオブジェクト種類、名前、保存方法を指定します。acSaveYesは設計変更を保存します。');
  if(mode===2)add(2300+i,'DoCmd・入出力','標準',`Excelファイルをテーブル「${table}」へ取り込むメソッドは？`,['DoCmd.TransferSpreadsheet','DoCmd.ImportExcel','DoCmd.TransferText','CurrentDb.ImportSheet'],'TransferSpreadsheetメソッドは、Accessと表計算ファイルの間でデータをインポートまたはエクスポートします。',`DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "${table}", path${i}, True`);
  if(mode===3)add(2300+i,'DoCmd・入出力','標準','アクションクエリのSQL文字列を直接実行するDoCmdメソッドは？',['DoCmd.RunSQL','DoCmd.ExecuteSQL','DoCmd.OpenSQL','Application.RunQuery'],'DoCmd.RunSQLはアクションクエリのSQLを実行します。確認メッセージを制御する場合は注意が必要です。',`DoCmd.RunSQL "DELETE FROM ${table} WHERE ID=${i}"`);
  if(mode===4)add(2300+i,'DoCmd・入出力','基礎','Accessのシステムメッセージ表示を一時的に無効にする記述は？',['DoCmd.SetWarnings False','Application.DisplayAlerts = False','DoCmd.Messages False','Warnings.Enabled = False'],'DoCmd.SetWarnings FalseでAccessのシステムメッセージを抑制します。処理後はTrueへ戻します。',`DoCmd.SetWarnings False\n' 処理 ${i}\nDoCmd.SetWarnings True`);
}

// 8. Access SQL（50問）
for(let i=1;i<=50;i++){
  const mode=i%5,amount=1000+i*10;
  if(mode===0)add(2350+i,'Access SQL','基礎',`金額が${amount}以上の行だけを抽出する句は？`,[`WHERE 金額 >= ${amount}`,`HAVING 金額 >= ${amount}`,`FILTER 金額 >= ${amount}`,`SELECT 金額 >= ${amount}`],'行を抽出する条件はWHERE句に記述します。HAVING句は集計後のグループを絞り込みます。',`SELECT * FROM T_売上 WHERE 金額 >= ${amount};`);
  if(mode===1)add(2350+i,'Access SQL','基礎',`売上日を新しい順に並べる指定は？`,['ORDER BY 売上日 DESC','SORT 売上日 DESC','ORDER 売上日 NEWEST','GROUP BY 売上日 DESC'],'ORDER BY句で列を指定し、DESCを付けると降順になります。',`SELECT * FROM T_売上${i} ORDER BY 売上日 DESC;`);
  if(mode===2)add(2350+i,'Access SQL','標準','両方のテーブルでキーが一致する行だけを結合する句は？',['INNER JOIN','LEFT JOIN','CROSS JOIN','UNION JOIN'],'INNER JOINは結合条件が両側で一致する行だけを返します。',`SELECT * FROM T_受注${i} AS O INNER JOIN T_顧客 AS C ON O.顧客ID=C.ID;`);
  if(mode===3)add(2350+i,'Access SQL','標準',`グループ化後、件数が${i+1}件以上のグループだけを残す句は？`,[`HAVING Count(*) >= ${i+1}`,`WHERE Count(*) >= ${i+1}`,`FILTER Count(*) >= ${i+1}`,`GROUP BY Count(*) >= ${i+1}`],'集計関数の結果を条件にする場合はGROUP BYの後のHAVING句を使用します。',`SELECT 顧客ID, Count(*) FROM T_受注 GROUP BY 顧客ID HAVING Count(*) >= ${i+1};`);
  if(mode===4)add(2350+i,'Access SQL','標準',`ID=${i}の行の状態を「完了」に変更するSQLは？`,[`UPDATE T_受注 SET 状態='完了' WHERE ID=${i};`,`MODIFY T_受注 状態='完了' WHERE ID=${i};`,`INSERT T_受注 SET 状態='完了' WHERE ID=${i};`,`ALTER T_受注 SET 状態='完了' WHERE ID=${i};`],'既存行の値を変更するにはUPDATE文でSET句とWHERE句を指定します。');
}

// 9. Access関数・ドメイン集計（50問）
for(let i=1;i<=50;i++){
  const mode=i%5,days=1+i%20,id=100+i;
  if(mode===0)add(2400+i,'関数・ドメイン集計','基礎','Nullを指定した代替値へ置き換えるAccess関数は？',['Nz','IsNull','NullToValue','Coalesce'],'Nz関数は式がNullなら第2引数の値を返し、Nullでなければ式の値を返します。',`total${i} = Nz(Me!金額, 0)`);
  if(mode===1)add(2400+i,'関数・ドメイン集計','標準',`テーブルからID=${id}の氏名を1件取得するドメイン関数は？`,['DLookup','DCount','DFirstOnly','DomainValue'],'DLookup関数は指定したテーブルまたはクエリから条件に合うフィールド値を1件返します。',`name${i} = DLookup("氏名", "T_顧客", "ID=${id}")`);
  if(mode===2)add(2400+i,'関数・ドメイン集計','基礎','条件に合うレコード数を数えるドメイン関数は？',['DCount','DLookup','DSum','CountDomain'],'DCount関数は指定したドメイン内で条件に合うレコード数を返します。',`n${i} = DCount("*", "T_受注", "担当ID=${id}")`);
  if(mode===3)add(2400+i,'関数・ドメイン集計','標準',`日付に${days}日を加える式は？`,[`DateAdd("d", ${days}, 対象日)`,`DatePlus(対象日, ${days})`,`DateDiff("d", ${days}, 対象日)`,`対象日.AddDays(${days})`],'DateAdd関数で間隔"d"、加算値、対象日を指定すると日単位で加算できます。');
  if(mode===4)add(2400+i,'関数・ドメイン集計','標準','2つの日付の間の日数を求める関数は？',['DateDiff','DateAdd','DatePart','DaysBetween'],'DateDiff関数は指定した間隔単位で2つの日付の差を返します。',`days${i} = DateDiff("d", 開始日, 終了日)`);
}

// 10. エラー処理・トランザクション（50問）
for(let i=1;i<=50;i++){
  const mode=i%5;
  if(mode===0)add(2450+i,'エラー・トランザクション','基礎',`エラー発生時にラベル Handler${i} へ移動する記述は？`,[`On Error GoTo Handler${i}`,`Try Handler${i}`,`On Error Resume Handler${i}`,`Catch Handler${i}`],'VBAではOn Error GoToとラベル名を使って、実行時エラーの処理位置を指定します。');
  if(mode===1)add(2450+i,'エラー・トランザクション','基礎','現在の実行時エラー番号を取得するプロパティは？',['Err.Number','Error.Code','Err.Code','Exception.Number'],'ErrオブジェクトのNumberプロパティに現在の実行時エラー番号が格納されます。',`Debug.Print Err.Number  '確認 ${i}`);
  if(mode===2)add(2450+i,'エラー・トランザクション','標準','既定のDAO Workspaceでトランザクションを開始するメソッドは？',['DBEngine.Workspaces(0).BeginTrans','CurrentDb.BeginTrans','DBEngine.StartTransaction','Workspace.Transaction = True'],'DAOのBeginTransメソッドはWorkspaceオブジェクトに対して呼び出し、トランザクションを開始します。',`DBEngine.Workspaces(0).BeginTrans  '処理 ${i}`);
  if(mode===3)add(2450+i,'エラー・トランザクション','標準','トランザクション内の変更を確定するメソッドは？',['CommitTrans','SaveTrans','EndTrans','ConfirmTransaction'],'CommitTransメソッドはトランザクション内で行った変更を確定します。',`DBEngine.Workspaces(0).CommitTrans  '確定 ${i}`);
  if(mode===4)add(2450+i,'エラー・トランザクション','標準','トランザクション内の変更を取り消すメソッドは？',['Rollback','CancelTrans','UndoTrans','RevertTransaction'],'Rollbackメソッドは未確定のトランザクションを取り消し、開始前の状態へ戻します。',`DBEngine.Workspaces(0).Rollback  '取消 ${i}`);
}

window.EXTRA_QUESTIONS=bank;
})();
