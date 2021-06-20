# UsingCharList
テキスト（複数可）に使われている文字一覧を取得します

# requirement
unity2017 or later
npoi2.5.1 or later

# features
* エクセル・テキスト内に存在する文字を全て抜き出し、一覧として出力します。  
* 複数のファイルやディレクトリに対して処理することが可能です。  
    * ディレクトリの場合、取り込みを行う拡張子を選択することもできます。  
* 解析可能フォーマットはテキスト、XLS(X) です。

# usage
1. NPOI から以下のファイルをダウンロードし、使用可能にします。  
  **BouncyCastle.Crypto.dll**  
  **ICSharpCode.SharpZipLib.dll**  
  **NPOI.dll**  
  **NPOI.OOXML.dll**  
  **NPOI.OpenXml4Net.dll**  
  **NPOI.OpenXmlFormats.dll**  
  
    * サンプルには NPOI が含まれています。万が一ランセンスに問題が生じた場合、サンプルから削除される可能性があります。

2. Assets/ に一覧化したいテキストを入れます。
3. Tools/Using Character List を実行します。  
  ![image](https://user-images.githubusercontent.com/85425896/122670127-49cb7880-d1fb-11eb-8cd7-9aa5abb110df.png)
    * 画像はサンプルプロジェクトのものです。
4. 「+」をクリック、import に「2」で作成したテキスト（ディレクトリでも構いません）、export に一覧ファイル名を設定します。
5. 「4」を複数作成することも出来ます。
6. 「Create」で一覧ファイルを作成します。  
  
    一覧ファイル  
    ![image](https://user-images.githubusercontent.com/85425896/122669393-0e7b7a80-d1f8-11eb-98a2-97a75eb0cc43.png)

# license
This sample project includes the work that is distributed in the Apache License 2.0 / MIT / MIT X11.  
このサンプルプロジェクトは、 Apache 2.0 / MIT / MIT X11 ライセンスで配布されている dll が含まれています。  

NPOI (Apache2.0): https://www.nuget.org/packages/NPOI/2.5.1/License  
SharpZLib (MIT): https://licenses.nuget.org/MIT  
Portable.BouncyCastle (MIT X11): https://www.bouncycastle.org/csharp/licence.html  
