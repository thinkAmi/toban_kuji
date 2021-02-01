toban_kuji
===

当番表を作るためのPythonスクリプトです。

### shuffle.py

ランダムに当番を決めるスクリプトです。

ファイル `users_for_shuffle.xlsx` を読み込んで使います。

　  

### optimization.py

組合せ最適化により当番を決めるスクリプトです。

制約条件はこちら

- 1当番あたり3人
- 一人あたり1回の当番
  - 複数回は不可
- その人が当番可能な時に割り当てる
  - 無理に当番をお願いすることはできない

ファイル `users_for_optimization.xlsx` を読み込んで使います。

　  

### make_data.py

テスト向けに、組合せ最適化用のファイル `users_for_optimization.xlsx` を作成するスクリプトです。

　  

## Blog

- [Python + PuLP + ortoolpy による組合せ最適化を使って、行事の当番表を作ってみた - メモ的な思考的な](https://thinkami.hatenablog.com/entry/2021/02/02/084026)
