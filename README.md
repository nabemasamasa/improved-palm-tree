# Hit and Blow

Gitの使い方や、作業の流れを理解するための課題用。  

## 使い方

1. GitHubで、画面右上の`Fork`をクリックし、自分のユーザー名を選択
    - Gitを知らない人向けに説明すると、Forkはコピーみたいなものと思っていい
1. コードレビュー用の設定をする
    1. GitHubで、HitAndBlowのリポジトリの`Setting`をクリック
    1. 左のメニューにある`Collaborators`をクリック
    1. `Search by username, full name or email address`のテキストボックスに  
    コードレビューしてもらいたいユーザーのIDを入れて`Add collaborator`をクリック
        - `TAKAHASHI Tatsuya`と`WATANABE Masahiro`が課題コードレビュー担当
1. Forkした自分のリポジトリを、PCに`Clone`する
    1. GitHubで、画面中央付近の緑色の`Code▼`をクリック
    1. `SSH`タブを選択し、`git@github.com:(自分のID)/Hit-and-Blow.git`をコピー
    1. GitBashで`git clone git@～(先程コピーしたもの)`を叩く
        - Bashでの貼り付け操作は`Shift`+`Insert`
1. 作業する
