# gas-webapp-sales-management-graph & ranking
- GASで作成したWebアプリ（営業管理グラフ & ランキング）

# GAS Webアプリ：営業管理グラフ & ランキング

## URL
### Webアプリ
https://script.google.com/macros/s/AKfycbxVZGo_RLr9WsijhPvD4MGqarYROEjginhEc0Af0tXYLHPmY_24zKxXSxPUTRGJxsVbaw/exec

### スプレッドシート（閲覧）
https://docs.google.com/spreadsheets/d/1qeceAOApZM0KVvGOGKjOlkub5KbYOabh6OOBBvyfHUc/edit?usp=sharing

## 概要
- 目的：営業マンの数字を「グラフ・ランキング」で表示し、競争心を芽生えさせモチベーションと生産性をUPさせるWebアプリ

## 機能
- 数字入力：営業マンにスプレッドシートを共有し入力（※下記15項目が入力する数字）
1. シフト
2. 稼働時間
3. 商談数
4. 必要売上金額
5. 売上金額
6. 必要受注数
7. 受注数
8. 必要コール数
9. 実コール数
10. 必要見込み数
11. 見込み数
12. 必要代表接触数
13. 代表接触数
14. 必要アポ数
15. アポ数

- 更新：1時間に1回数字を更新するトリガー設定、Webアプリの右上にある「更新」ボタンをクリックするとリアルタイムで反映

- グラフ：「当日・月間」ごとに「コール数・アポ数・受注数」をグラフで閲覧、「アポ率・アポ生産性・売上金額」を数字で閲覧

- ランキング：「当日・月間」ごとに「コール数・アポ数 / アポ率・アポ生産性・受注数・売上金額」をランキングで閲覧

- レスポンシブ対応：スマホ画面で見やすいようにデザインを調整（グラフ画面の横一列を2名で表示、ランキング画面の項目を縦一列で表示）

## 実装予定
- 数字入力の簡易化：スプレッドシートに直接入力ではなく、Googleフォームの回答の入力に変更（回答した数字がスプレッドシートに自動で反映される）

## 技術
- Google Apps Script

- Google スプレッドシート
 
- HTML / CSS / JavaScript

## 画面イメージ

### PC画面
- グラフ画面
<img width="1656" height="796" alt="営業管理グラフ_PC画面" src="https://github.com/user-attachments/assets/7923ef9c-3ce1-47c9-b147-78cda05608b4" />

- ランキング画面
<img width="1663" height="504" alt="営業管理ランキング_PC画面" src="https://github.com/user-attachments/assets/1908b9f1-018f-407e-b2ef-1b329584bb12" />

### スマホ画面
- グラフ画面
<img width="380" height="597" alt="営業管理グラフ_スマホ画面" src="https://github.com/user-attachments/assets/931d4063-9630-487f-9019-f4a9b4e1a6b1" />
  
- ランキング画面
<img width="254" height="661" alt="営業管理ランキング_スマホ画面" src="https://github.com/user-attachments/assets/af478078-be9d-4df2-a316-59c56966b635" />
