# invoice_management_system_vba
# 納品書・請求書発行システム (Microsoft Access)

中小規模事業者向けの、Microsoft Accessベースの納品書・請求書発行システムです。
入力の効率化（オートコンプリート）と、過去データの履歴保持（非正規化）を特徴としており、B5用紙への印刷に最適化されています。

## 📋 機能概要

* **入力支援機能**: 顧客名や商品名の「読み」を入力するだけで候補を表示し、住所や単価を自動補完します。
* **履歴データの保持**: マスタデータ（住所や単価）が将来変更されても、過去の伝票データ（発行時点の情報）は維持される設計です。
* **B5印刷対応**: B5用紙（縦）に合わせて、納品書と請求書を連続で出力します。

## 🛠 技術スタック / 環境

* **Platform**: Microsoft Access (Office 365 / 2019以降推奨)
* **Language**: VBA (少量のイベント制御), Access Macro
* **File Format**: `.accdb`

---

## 💾 データベース構造 (Schema & ER Diagram)

本システムは、典型的な「ヘッダー・明細」形式のリレーショナルデータベース構造を持ちつつ、**履歴保持のためにあえて正規化を崩したフィールド（非正規化フィールド）** を持っています。

### ER図 (Mermaid)

```mermaid
erDiagram
    T_SALES ||--|{ T_SALES_DETAIL : "1対多 (One-to-Many)"
    T_CUSTOMER ||--|{ T_SALES : "参照 (Lookup)"
    T_ITEM ||--|{ T_SALES_DETAIL : "参照 (Lookup)"

    T_SALES {
        long 売上ID PK "AutoNumber"
        long 顧客ID FK "参照のみ"
        date 売上日付
        string 納品先名_保存用 "履歴保持"
        string 住所情報_保存用 "履歴保持"
    }

    T_SALES_DETAIL {
        long 売上明細ID PK "AutoNumber"
        long 売上ID FK
        long 商品ID FK "参照のみ"
        string 商品名_保存用 "履歴保持"
        currency 単価_保存用 "履歴保持"
        int 数量
    }
