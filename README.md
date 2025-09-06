## 專案說明

# FCU Class

逢甲大學選課自動化工具，支援自動登入、驗證碼辨識、課程加選，並提供簡易 GUI 操作介面。

## 功能特色

-   自動化登入逢甲選課系統
-   驗證碼自動辨識（ddddocr）
-   支援多科目加選
-   GUI 介面（PySide6）
-   支援 cookies 快速重登入
-   設定檔管理（config.ini）

## 安裝方式

1. 安裝 Python 3.11 以上版本
2. 安裝依賴套件：
    ```bash
    pip install -r requirements.txt
    ```
    或使用 uv：
    ```bash
    uv pip install -r requirements.txt
    ```

## 使用方法

1. 執行 GUI 介面：
    ```bash
    uv run main.py
    ```
    或
    ```bash
    python main.py
    ```
2. 依照介面操作，開始自動選課

## 主要檔案說明

-   `main.py`：GUI 介面
-   `course.py`：核心選課流程
-   `config.ini`：使用者設定檔
-   `requirements.txt`：依賴套件清單

# 注意事項

僅供學術研究，使用完成請刪除檔案

# 版權

此專案的版權規範採用 MIT License - 至 LICENSE 查看更多相關聲明
