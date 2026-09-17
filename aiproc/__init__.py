"""一覧表の文章列のAI整形（keep モード）。

- prompts: 送る messages の組み立て（固定ルール＋版ごとの自動生成部分＋行ごとのデータ）
- verify: AI出力の原文照合（項目単位の合否）
- cache: llm_calls（生の応答のキャッシュ。即時コミット）
- items: ai_items（行×段の状態とハッシュ、古くなった判定）
- runner: AIジョブ本体と試し実行
- custom: custom 段（短い文 / 選択肢1つ）
- estimate: 所要時間・トークンの見積もり
"""
