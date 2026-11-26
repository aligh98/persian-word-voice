---
name: persian-word-voice
description: Voice a whole Persian (Farsi) vocabulary list or word database with ElevenLabs so each word or expression gets a native, correct, natural pronunciation clip. Built for language-learning apps (flashcards, dictionaries, vocabulary lessons). It cleans Arabic look-alike letters, finds homographs (مهر، ملک، کشتی...), adds vowel marks only where a word has more than one reading, generates on eleven_v4, builds a listening page for ear review, and re-voices reported words with a meaning-matched carrier sentence. Use it whenever someone wants Persian word or phrase audio, pronunciation files, "صداگذاری کلمات", «تلفظ لغات», «فایل صوتی برای هر کلمه», a Persian TTS batch from an .xlsx/.csv word list, or asks why ElevenLabs mispronounces a Persian word, even if they never mention this skill.
---

# Persian word voice

This skill turns a list of Persian words and expressions into one audio file per entry, with the pronunciation a native teacher would use. The method is validated by native-speaker listening tests on difficult verbs, homographs, Arabic-origin expressions and one-syllable words. Each rule below says why it exists. Follow the rules, and do not "improve" them with techniques the tests rejected.

Talk to the user in their language (usually Persian). Keep this file's English for yourself.

## Tool

All work goes through one script and one table:

```bash
python <skill>/scripts/fa_words.py <command> ...
```

The table is `<work>/words.tsv`, one row per word. Its columns:

| Column | Meaning |
|---|---|
| `word` | Clean display form |
| `meaning` | Optional translation or meaning column |
| `kind` | `single`, `homograph` or `review` |
| `spoken` | Text sent to the TTS |
| `method` | `base` or `carrier` |
| `sentence` | Carrier sentence |
| `status` | `new`, `done`, `flagged` or `ok` |
| `note`, `made_from` | Notes, and the text the current audio was made from |

Change rows with the `set` command, not by hand-editing the file, so the checks run.

Needs: Python 3.9+, `pip install openpyxl` for .xlsx, and `ffmpeg` (only for carrier cuts). The key comes from the environment variable `ELEVENLABS_API_KEY`. Never write the key into a file, a command you show, or the skill. If the user pastes a key in chat, tell them to set it as an environment variable, and to rotate it later.

## Fixed settings (tested, do not change unless asked)

- **Model `eleven_v4`, `language_code: fa`.** Persian is supported only on the v4 and v3 families. `multilingual_v2`, flash and turbo do not list Persian. `eleven_v3` mangled short isolated words («گرم» came out «یا ارم», «درآوردن» came out «دارا بردان»). Do not use v3.
- **Voice: the user's choice.** Any ElevenLabs voice works: a Persian voice from the Voice Library, the user's own cloned voice, or one made with Voice Design. Ask for the Voice ID before the first run. If the user has none, suggest browsing Voices with the language filter set to Persian and copying the ID. Otherwise use the suggested default, Sara (`RBEKbcQ62sKk2tsmPsAU`, "Sara - Neutral Persian Female"). The voice is a matter of taste, not a fix: right and wrong readings were the same across every Persian voice tested. `voices` lists the Persian voices in the account. Pass the ID once with `generate --voice ID`. It is saved in `config.json`, so the whole project stays in one voice. Switching later needs `--force`, which regenerates everything. A Library voice must first be added to "My Voices".
- **Seed 42, stability 0.5, similarity 0.8.** Different seeds sounded the same by ear, so the seed is only there to make reruns repeatable. Re-rolling seeds is not a fix for a wrong word. `eleven_v4` ignores `speed` and `style`.
- **Judge by ear, not by speech-to-text.** Scribe writes back the dictionary spelling. It cannot tell «مَرد» from «مُرد», and it gave two answers for one file. Do not use it to approve words.

## Workflow

### 1. Prepare

```bash
python fa_words.py prepare words.xlsx --work work/ [--column کلمه] [--meaning-column معنی]
```

The command reads every sheet, skips title and header rows (like «مصدرها» / «کلمه»), drops duplicates and cleans Unicode:
- Arabic «ي ى ك» become Persian «ی ک»;
- tatweel is removed;
- spaces and ZWNJ are tidied.

The Unicode step matters. Arabic code points pull the model towards an Arabic accent, and databases copied from the web are full of them.

The command reports the word count, how many were cleaned, and the credits a base run would cost (1 credit per character). Tell the user these numbers.

If the source has a translation or meaning column, pass it with `--meaning-column`. It is the best evidence for choosing between the readings of a homograph.

### 2. Decide the spelling of each word (your judgment)

This is the core of the skill. The rule, validated in listening tests:

| Case | What to send | Why |
|---|---|---|
| One common reading, even if hard or Arabic-origin (ایران، من‌حیث‌المجموع، حتی‌الامکان، دست‌وپاچلفتی، گلنگدن) | The word as written, **no marks** | The model knows the real spelling and says it with natural feeling. Marks add nothing and can hurt. |
| More than one common reading (مهر، ملک، کشتی، گزیدن، گرم، مرد) | **Minimal marks**: only the letter that picks the reading («مُهر»، «مَلِک»، «کَشتی»، «گُزیدن») | Without the mark the model picks the reading on its own. Context did not save «مهر»: it was read «مِهر» even in a sentence about a seal. |
| Compounds with ezafe (آب معدنی، اتاق خواب، صاحب خانه، شب یلدا) | Plain, **no ezafe kasra** | The model adds the ezafe on its own. In listening tests the kasra versions («اتاقِ خواب»، «صاحبِ خانه») sounded worse than plain. A homograph inside a compound still gets its minimal mark: «حُسن نیت», since plain «حسن» was read as the name Hasan. |
| Anything | **Never respell** («مِن‌هِیسُل‌مَجمو» style) and never mark every letter | A respelled word is a string the model does not know. It comes out flat and lifeless. Full marks make Persian sound Arabic. |

How to apply it:
- `list work/ --kind homograph` shows the flagged entries, with their readings in `note`. An entry is flagged when it is in `data/homographs.tsv`, when one of its words is («کشتی گرفتن» contains «کشتی»), or when the same spelling is listed twice with different meanings. In that last case each row keeps its own reading: «مُهر» for seal, «مِهر» for love.
  - For each one, choose the reading from the `meaning` column, the sheet name (a verb list means «مُرد», not «مَرد»), or the app's context.
  - Then run `set work/ ID --spoken "مُهر"`.
  - If the meaning is truly unknown, ask the user. Do not guess. A wrong homograph teaches the learner the wrong word.
- The list is not complete. Read through the `single` words too, and treat any other word whose reading depends on its meaning the same way.
  - Batch this: read the table in chunks of 200 or so and set only the words that need it.
  - Most words need nothing.
- Example sentences work the same way as words: one row per sentence. Inside a sentence, function words such as «در»، «کرد»، «کند» get flagged too, but the grammar already settles them, so leave them unmarked. Mark only content words whose reading the sentence leaves open («پایش مُهر زد»، «سوار کَشتی شدند»). For sentences, keep the default `--names id`.
- `kind=review` means the source already had marks. Keep only the marks that pick between readings, and drop the rest.
- `set` refuses a `--spoken` that changes letters. That guard is the "never respell" rule.

### 3. Generate (base)

```bash
python fa_words.py generate work/ --dry-run      # count and credits first
python fa_words.py generate work/
```

"Base" means the word alone, as `spoken`. It was the best method in every round. It is resumable: words already made from their current text are skipped, so a run stopped by credits or the network can just be started again. Before a large run, check the account balance with `estimate work/`. For a big list, voice 10–20 mixed words first and let the user listen before spending credits on everything.

### 4. Ear review

```bash
python fa_words.py review work/ [--status done]
```

This writes `work/listen.html`: one row per word with a player, a "wrong/weak" checkbox and a note field. The "ساخت فهرست" button turns the ticks into a short list the user pastes back to you.

Ask the user to open the page and listen. For each reported word, run `set work/ ID --status flagged --note "..."`. Mark the approved ones `ok` when the user says the rest are fine (`set` per ID, or ask them to say "the rest are OK").

### 5. Fix reported words: in this order

1. **Spelling first** (free). Is it a homograph that went out unmarked, or did it get a mark it did not need? Fix `spoken` by the rule in step 2 and regenerate only those IDs with `generate work/ --ids 12,40`.
2. **Last resort: a dynamic carrier sentence.** Use it only when the user chooses it, for words that are still wrong or weak. See below.

### 6. The carrier fallback (only on the user's choice)

**What it is:** a short, natural sentence written for that word's meaning and part of speech. The word comes last, after a colon and in quotes. The whole sentence is voiced, then the word is cut out at the silence before it. The context gives the model the right melody and certainty, and the cut keeps only the word.

**Offer, don't impose.** Tell the user the method exists and that it costs more ElevenLabs credits: the whole sentence is billed, about 40–80 characters instead of about 6, so roughly ten times a base word. Use it only if they say yes, and only on the words they pick.

Writing the sentence:
- Work out what the word means and how it is used (verb, noun, adjective, idiom, adverb), then write one sentence that leads to it naturally:
  - a definition («به کسی که … می‌گن»);
  - a scene («کارمند نامه رو امضا کرد و پاش زد»);
  - a usage frame («در زبان ادبی می‌گیم»).
- End exactly with `: "X"`, where X is the `spoken` form (minimal marks kept for homographs). The word must not appear earlier in the sentence. `set` checks both.
- Use plain spoken Persian, 40–80 characters, no audio tags, and no other homographs near the end.
- Do not use a fixed frame like «معلم زبان تلفظ کرد: "X"». Listening tests rejected it: a generic frame gives a generic, flat reading. The sentence must fit the word.

Examples that passed the listening test:

| Word | Sentence |
|---|---|
| دست‌وپاچلفتی | `به کسی که همه‌چیز از دستش می‌افته و هر کاری رو خراب می‌کنه، می‌گن: "دست‌وپاچلفتی"` |
| من‌حیث‌المجموع | `وقتی همه‌ی جوانب رو کنار هم بذاریم و بخوایم نتیجه بگیریم، می‌گیم: "من‌حیث‌المجموع"` |
| گلنگدن | `سرباز برای مسلح کردن تفنگ، اهرم کنارش رو کشید؛ اسم اون اهرم هست: "گلنگدن"` |
| مَلِک | `توی داستان‌های قدیمی، به پادشاه می‌گفتن: "مَلِک"` |
| گُزیدن | `وقتی از بین چند گزینه بهترین رو انتخاب می‌کنیم، در زبان ادبی می‌گیم: "گُزیدن"` |

Steps:
1. Write the sentences.
2. **By default, show the user only 3 of them as examples,** with the credit cost from `estimate work/ --ids ...`, and wait for their go-ahead.
3. Then run `set work/ ID --method carrier --sentence '...'` for each word, followed by `generate work/ --ids ...` and `review work/ --ids ...`.
4. The full sentence audio is kept in `audio/carrier_full/`, in case a cut sounds off.

**Optional extras: offer them once, and do them only if the user accepts.** They are not part of the default run:
- a full review of every carrier sentence before generating, instead of 3 examples;
- a detailed cost breakdown;
- keeping `words.tsv` as the long-term lexicon (method, spelling, sentence and status per word), so a future rebuild reuses approved choices without retesting.

### 7. Export

```bash
python fa_words.py export work/ --out final/ [--names id|word] [--loudness [DB]]
```

`--loudness` brings every clip to the same average level, -20 dBFS by default, without clipping. Words cut from a carrier sentence can come out a little louder or softer than base words, and a flashcard app should not jump in volume from card to card. Offer it once. It needs ffmpeg.

This copies one mp3 per word and writes `manifest.csv` (id, word, meaning, file, method, spoken, status) for the app's import. `--names id` (default) gives ASCII-safe names like `0042.mp3`. `--names word` uses the Persian word. The command warns about words with no audio yet and words still flagged.

## When something goes wrong

| Symptom | What to do |
|---|---|
| 401 | Wrong key, or the key lacks text-to-speech permission. |
| Voice not found | Add Sara (or the chosen voice) to "My Voices" in ElevenLabs, or pass `--voice`. |
| Out of credits | The run stops and saves progress. Top up, then run `generate` again: done words are skipped. |
| A carrier cut clips the start of the word or keeps a bit of the sentence | Listen to `audio/carrier_full/ID.mp3`. Make the colon pause stronger (end the lead-in with «؛» or «:» right before the quote), or rewrite the lead-in, then regenerate that ID. |
| The user reports a word sounds Arabic | Check for leftover marks or Arabic code points. Remove marks that are not needed to pick a reading. |
