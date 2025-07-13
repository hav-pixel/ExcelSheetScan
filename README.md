# はじめに

このツールは、仕事で Excel ドキュメント内のテキストを効率的に検索・置換するために作成しました。  

**セルやシェイプの文字列**を対象に、複数ファイルを横断的に検索できます。

※検索・置換機能は **実ファイルを変更する可能性があるため注意**してください。

---

## ✅ 主な機能

- 複数の Excel ファイル内の **セル／シェイプの文字列を検索**
- キーワードの **完全一致／曖昧一致（FUZZY）** 対応
- **検索＋置換**（オプション機能）
- 検索結果を Excel に貼り付け、**HYPERLINK関数等使いジャンプ可能**

---

## 🔧 ビルド手順

```bash
.\gradlew clean installDist
````

---

## 🚀 実行方法

```bash
build\install\SearchDocs\bin\SearchDocs.bat testData\ "Jakarta" FUZZY
```

出力例:

```
以下の条件でgrep検索を実行します。
検索対象フォルダ：testData\
検索文字列：Jakarta
検索方法：FUZZY

2       0       -1      C:\...\jakartaEE-HSSF.xls      Sheet2  D18     これは D18 Jakarta です。
3       2       8       C:\...\jakartaEE.xlsx          Sheet2  B89     Jakarta EE 9.1 → Jakarta EE 10
...（略）
```

---

## 🧪 Excel からリンクして確認する

`result.txt` に出力して、Excel からリンクでジャンプ可能にできます。

```bash
.\gradlew build install
.\build\install\SearchDocs\bin\SearchDocs.bat .\testData\ "Jakarta" "FUZZY" > result.txt
```

* `result.txt` を Excel に貼り付け
* 下記のような数式で HYPERLINK を作成できます：

```excel
=HYPERLINK(D1 & "#" & E1 & "!" & F1)
```

例：

```
C:\...\SearchDocs\testData\dir01\テスト用D01F01.xlsx#Sheet1!E37
```

---

## 🐞 SLF4Jの初期化メッセージについて

ログ出力に以下の情報メッセージが出ることがありますが、**警告ではありません**。

```
SLF4J(I): Connected with provider of type [ch.qos.logback.classic.spi.LogbackServiceProvider]
```

消したい場合は、`logback.xml` でログレベルを `WARN` 以上に設定してください：

```xml
<configuration>
  <root level="WARN">
    <appender-ref ref="CONSOLE" />
  </root>
</configuration>
```

---

## 📁 フォルダ構成例

```
SearchDocs/
├─ build/
├─ docker/
├─ testData/
├─ src/
├─ README.md
└─ settings.gradle
```
