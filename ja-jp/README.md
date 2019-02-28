# VBScript for adding Japan new Holidays to Outlook

2019 年以降、日本の祝日に追加や変更があります。

メイン ストリームの製品 (本投稿時点では Outlook 2016 / 2019 および Office 365 ProPlus) の Outlook をご利用の場合、祝日の追加 / 変更は更新プログラムによって対応しております。  
しかし、Outlook 2013 以前のバージョンは延長サポートの製品であるため、対応する更新プログラムをリリースする予定はありません。  
新しい祝日を追加するスクリプトを作成しましたので、こちらでの対応をご検討ください。

このスクリプトを実行すると、2019 年以降の祝日が Outlook の既定の予定表に追加されます。  
このスクリプトでは、既に追加されている祝日は追加せず、追加されていない祝日を追加するため、既に祝日が追加されている場合でも、まだ祝日が追加されていない場合でも、実行できます。

## 実行方法

1. [AddHolidays.zip をダウンロード](https://github.com/Microsoft/AddOutlookJapanHoliday/releases)し、展開します。
2. 新しい祝日を追加したい Outlook 2010 / 2013 の環境で、AddHolidays.vbs をダブルクリックします。
3. Outlook を起動して新しい祝日が追加されたことを確認します。

## フィードバック

スクリプトに関するフィードバックは [Issues](https://github.com/Microsoft/AddOutlookJapanHoliday/issues) に投稿してください。日本語でも構いません。

本プロジェクトへの参加に関しては、[英語版 README の Contributing セクション](https://github.com/Microsoft/AddOutlookJapanHoliday#contributing)をご参照ください。