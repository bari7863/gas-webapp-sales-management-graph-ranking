# gas-webapp-sales-dx
- GASで作成したWebアプリ

# GAS Webアプリ①：営業管理グラフ & ランキング

## URL
- Webアプリ：https://script.google.com/macros/s/AKfycbxwAN3VyIEUyZglxmqWbWABV75NYjhUr_CO--rxIjWEdx5hA6laBBc8pePhDpFE-WFH/exec
- スプレッドシート（閲覧）：https://docs.google.com/spreadsheets/d/1sbtPfK7p7SmPbTWP1zzJJtjJryMFIwNmCHO073lvA1I/edit?usp=sharing

## 概要
- 目的：営業マンの数字をグラフ・ランキングで表示し、競争心を芽生えされるためのWebアプリ

## 機能
- 数字入力：営業マンそれぞれに各スプレッドシートを共有し、入力してもらう
- 更新：1時間に1回数字を更新するトリガー設定、Webアプリの右上にある「更新」ボタンをクリックするとリアルタイムで反映
- グラフ：「当日・月間」ごとに「アポ・商談数 / 商談化率・受注数」をグラフで閲覧
- ランキング：「当日・月間」ごとに「アポ・商談数 / 商談化率・受注数」をランキングで閲覧

## 実装予定
- 数字入力の簡易化：スプレッドシートに直接入力ではなく、Googleフォームの回答の入力に変更（回答した数字がスプレッドシートに自動で反映される）

## 技術
- Google Apps Script
- Google スプレッドシート
- HTML/CSS/JavaScript

## 画面イメージ
- グラフ画面
 <img width="1665" height="671" alt="営業管理グラフ" src="https://github.com/user-attachments/assets/87127c15-5c56-44b3-86ec-fa865c391c62" />
  
- ランキング画面
<img width="1669" height="455" alt="営業管理ランキング" src="https://github.com/user-attachments/assets/f37ae2ff-6ed4-4c3d-be49-0d17c0d9da47" />
