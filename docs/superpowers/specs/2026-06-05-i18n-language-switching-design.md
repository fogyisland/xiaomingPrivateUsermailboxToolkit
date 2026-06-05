# i18n Language Switching — Design

**Date:** 2026-06-05
**Status:** Approved
**Author:** Claude (brainstorming with user)

## Goal

Add Chinese / English language switching to LiteEMLTOPST. Users pick a language in the existing PreferencesWindow (常规设置) and the entire UI is translated after restart. All UI text lives in external XML files under `Language/`, so translators can edit them without recompiling.

## Scope

In scope:
- 2 supported languages: `zh-CN` (default) and `en-US`
- New `Language/` directory with `zh-cn.xml` and `en-us.xml`
- `LocalizationManager` service for loading + lookup
- `LocExtension` markup extension for XAML binding
- New "🌐 语言" Tab in `PreferencesWindow`
- All UI text in 6 XAML windows and any user-facing strings in code-behind

Out of scope:
- More than 2 languages
- Hot-reload of language files (restart required)
- Right-to-left layouts
- Number/date/currency formatting (text-only translation)

## Architecture

```
Language/
├── zh-cn.xml              # Chinese (default + fallback)
└── en-us.xml              # English

Services/
└── LocalizationManager.cs # Singleton loader + cache + lookup

Markup/
└── LocExtension.cs        # WPF MarkupExtension for XAML

PreferencesWindow.xaml    # Add "🌐 语言" Tab
```

`usersettings.json` schema gains one field: `"Language": "zh-CN" | "en-US"`.

## Components

### `LocalizationManager` (singleton)

Static class in `Services/LocalizationManager.cs`.

- `Initialize(string culture)` — called from `App.OnStartup`. Loads `Language/{culture}.xml` into `Dictionary<string, string>`. Always also preloads `Language/zh-cn.xml` as fallback.
- `GetString(string key)` — lookup order: current culture → zh-CN fallback → return `"[key]"` (bracketed so missing keys are obvious in the UI).
- `GetString(string key, params object[] args)` — supports `string.Format` placeholders.
- `CurrentLanguage` — current culture string (`"zh-CN"` or `"en-US"`).
- `AvailableLanguages` — hardcoded list: `("zh-CN", ...)` and `("en-US", ...)`, where the display name comes from `Language_DisplayName_<culture>` keys in the current culture's XML (e.g. zh-CN: "中文（简体）", en-US: "English"). In English mode, the zh-CN display name falls back to the key from the zh-CN fallback file (still showing "中文（简体）" in the picker — accepted limitation; users will recognize 中文).
- Thread-safe (concurrent reads after init, init happens once on UI thread).
- Never throws — all file/parse errors are logged and fall back to defaults.

### `LocExtension` (MarkupExtension)

Class in `Markup/LocExtension.cs` in namespace `MailConvertPrivateUser.Markup`.

```xml
xmlns:loc="clr-namespace:MailConvertPrivateUser.Markup"
...
<TextBlock Text="{loc:Loc Preferences_Title}"/>
```

- `Key` property
- `ProvideValue` returns `LocalizationManager.GetString(Key)`
- No live refresh — values are baked at XAML load time (matches our "restart to apply" choice)

### XML file format

```xml
<?xml version="1.0" encoding="utf-8"?>
<resources>
  <string name="Preferences_Title">常规设置</string>
  <string name="Tab_EmlToPst">📧 EML转PST</string>
  ...
</resources>
```

Key naming convention: `{Window}_{Element}_{Purpose}` (e.g. `Main_BtnConvert_Click`, `Preferences_Tab_Language`). Keys never contain spaces or special chars.

### `PreferencesWindow` new Tab

Add a 5th Tab with header bound to `{loc:Loc Preferences_Tab_Language}` (key resolves to "🌐 语言" in Chinese, "🌐 Language" in English). Contents:
- `ComboBox` bound to `AvailableLanguages` (display = friendly name from key `Language_DisplayName_zh-CN` / `Language_DisplayName_en-US`, value = culture code)
- Currently selected item matches `LocalizationManager.CurrentLanguage`
- The existing "保存" button at the bottom of PreferencesWindow writes the chosen culture to `usersettings.json["Language"]` alongside other settings
- On save, show `MessageBox` with text from key `Preferences_Language_RestartPrompt` (zh: "语言已切换为 {0}，请重启应用后生效。", en: "Language switched to {0}. Please restart the application to take effect.")

**Filename vs culture-code note:** XML files use lowercase per filesystem convention (`zh-cn.xml`, `en-us.xml`); culture codes in `usersettings.json` and `AvailableLanguages` use Microsoft's mixed-case convention (`zh-CN`, `en-US`). The mapper is `culture.ToLowerInvariant()` when reading the file.

## Data Flow

### Startup

1. `App.xaml.cs OnStartup` reads `usersettings.json["Language"]`; default `"zh-CN"` if missing/invalid
2. Calls `LocalizationManager.Initialize(culture)`
3. `LocalizationManager` reads `Language/{culture}.xml` + `Language/zh-cn.xml`, populates dictionaries
4. XAML windows are loaded; `{loc:Loc ...}` markup extensions call `GetString` and return current language text
5. Code-behind uses `LocalizationManager.GetString("Key")` for dynamic strings (MessageBoxes, etc.)

### Switching

1. User opens PreferencesWindow → "🌐 语言" Tab → selects "English" → clicks "保存"
2. `usersettings.json["Language"]` is updated to `"en-US"`
3. `MessageBox` informs user that restart is required
4. `PreferencesWindow` closes
5. Current session still shows old language (no half-translated UI)
6. User manually restarts the app → startup loads `en-us.xml` → UI displays English

### Invariants

- `LocalizationManager`'s in-memory dictionaries never change during a session
- Any code path that needs a localized string calls `GetString` synchronously (no caching in callers)
- The `[key]` bracketed fallback makes untranslated strings visually obvious so they can be reported

## Error Handling

| Scenario | Behavior |
|----------|----------|
| `Language/{culture}.xml` not found | Log warning, use hardcoded fallback map (subset of most common strings); UI degrades to English-like behavior |
| XML parse error | Same as above |
| `usersettings.json` missing `Language` field | Default to `"zh-CN"` |
| Missing key in current culture | Look up in `zh-cn.xml`; if still missing, return `"[key]"` |
| `usersettings.json["Language"]` set to unsupported value | Treat as missing → default `"zh-CN"` |
| Runtime edit of XML | No effect; restart required |
| `LocExtension` used in non-UI context | `GetString` is thread-safe pure dict lookup; safe |

Logging via existing Serilog to `logs/emltopst-.log`:
- `Localization initialized for {culture} ({loaded} keys)`
- `Localization: missing key {key} (culture: {culture})`
- `Localization: failed to load {path}: {error}`

## Translation Process

1. Sweep all XAML files and key code-behind for user-facing Chinese strings
2. Build a `key → 中文` mapping table
3. Write `Language/zh-cn.xml` with all Chinese values
4. Translate each value to English; write `Language/en-us.xml`
5. User reviews the translation table before commit

Approximate scope (from initial exploration): 6 XAML windows, ~200-300 unique strings total. This is a one-time effort; future strings added by the user follow the same `loc:Loc` pattern.

## Testing

No automated test infrastructure exists. Testing is manual smoke tests:

- [ ] First launch → UI in Chinese
- [ ] PreferencesWindow shows "🌐 语言" Tab with Chinese as current
- [ ] Switch to English → save → "请重启" prompt appears
- [ ] Restart app → UI in English
- [ ] Delete `Language/zh-cn.xml` → restart → no crash, app still usable
- [ ] Corrupt `Language/en-us.xml` (bad XML) → restart → no crash, falls back gracefully
- [ ] Set `usersettings.json["Language"] = "fr-FR"` → restart → falls back to Chinese
- [ ] Check all 6 windows: MainWindow, PreferencesWindow, ImapAccountEditWindow, O365AccountEditWindow, ContactMappingDialog, RegistrationWindow
- [ ] MessageBox strings (save success, error, registration, etc.) all translated

## Implementation Order

1. Create `Language/zh-cn.xml` and `Language/en-us.xml` with all extracted strings
2. Implement `Services/LocalizationManager.cs`
3. Implement `Markup/LocExtension.cs`
4. Wire `App.OnStartup` to call `LocalizationManager.Initialize`
5. Refactor each XAML window to use `{loc:Loc Key}` (start with PreferencesWindow since it owns the new Tab)
6. Refactor code-behind to use `LocalizationManager.GetString` for MessageBoxes (note: `App.xaml.cs` DispatcherUnhandledException handler at line 41-42 also has hardcoded "错误" string — must be translated)
7. Add the "🌐 语言" Tab in PreferencesWindow
8. Update `usersettings.json` save/load to include `Language` field
9. Build, smoke test, fix any untranslated strings
10. Commit + version bump to v1.0.0.7

## Open Questions

None — all design decisions confirmed during brainstorming:
- Languages: zh-CN, en-US
- Switch location: PreferencesWindow "🌐 语言" Tab
- Effect: requires restart
- Implementation: custom LocalizationManager + LocExtension
- XML files: create from scratch
