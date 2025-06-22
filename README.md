# KENKYU1
LMJと投球の床反力を分析します．

## 使い方

1. 依存パッケージをインストールします。
   ```bash
   pip install -r requirements.txt
   ```
2. GUI を起動します。
   ```bash
   python main_KENKYU1.py
   ```

### CSV フォーマット
解析に用いる CSV は 5 行目がヘッダー行となっている必要があります。
最低限 `Time`, `Force.Fy.1`, `Force.Fz.2` といった列が含まれていることを想定しています。

### 分析モード
- **LMJ**: `Force.Fy.1` 列を用いた垂直跳びの分析を行います。
- **投球**: 軸足の `Force.Fy.1` と踏み込み足の `Force.Fz.2` を用いた投球動作の分析を行います。
