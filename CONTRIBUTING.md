# Работа с исходниками и релизами

## Правило расположения

Не использовать абсолютные пути к каталогу проекта. Корень репозитория всегда определяется через Git:

```powershell
$repoRoot = (git rev-parse --show-toplevel).Trim()
Set-Location $repoRoot
```

Все пути внутри команд строятся от `$repoRoot` через `Join-Path`.

Проверка подключения к GitHub:

```powershell
git remote -v
git branch --show-current
```

Основная ветка: `main`.

## Начало работы

Из каталога репозитория выполнить:

```powershell
$repoRoot = (git rev-parse --show-toplevel).Trim()
Set-Location $repoRoot
git status --short
git branch --show-current
git remote -v
git log -3 --oneline --decorate
```

Незакоммиченные изменения считаются пользовательскими. Не использовать `git reset --hard`, `git clean` и массовое восстановление файлов.

## Сборка

Основная сборка выполняется через `build_one_exe.bat`. Скрипт использует собственный каталог расположения через `%~dp0`, поэтому не зависит от имени и места расположения репозитория.

Для ручной сборки из PowerShell:

```powershell
$repoRoot = (git rev-parse --show-toplevel).Trim()
$python = Join-Path $repoRoot '.build311fresh\Scripts\python.exe'
$spec = Join-Path $repoRoot 'PROJECTZ GS1 DataMatrix 4.0.7.5.spec'

& $python -m PyInstaller --noconfirm --clean $spec
```

Текущая сборочная среда — `.build311fresh`. Каталоги `build/`, `dist/`, `tmp/` и виртуальные окружения не публикуются в Git.

В spec-файле нельзя добавлять путь к конкретному профилю пользователя. Путь к Ghostscript должен определяться через `PATH`, переменную окружения `GS_ROOT` или стандартные системные каталоги Windows.

## Изменение версии

При выпуске новой версии синхронно проверить:

- `APP_VERSION` в `app.py`;
- версии в `version_info.txt`;
- имя и `name` в новом spec-файле;
- номер версии в README и build-скриптах.

Старые EXE, build-кэши и временные QA-файлы не являются исходниками релиза.

## Проверки перед commit

```powershell
$repoRoot = (git rev-parse --show-toplevel).Trim()
Set-Location $repoRoot

& .\.build311fresh\Scripts\python.exe -m PyInstaller --version
& .\.build311fresh\Scripts\python.exe -c "import app; code=app.normalize_code(app.DEFAULT_CODE); ok, messages, _=app.validate_code(code, strict_gtin=False); assert ok, messages; symbol=app.generate_symbol(code, scale=2); app.decode_and_verify(symbol, app.build_expected_payload(code)); print('source self-test OK')"

git diff --check
git status --short
```

GUI-сценарий после изменений:

1. Запустить актуальный EXE.
2. Нажать «Создать и проверить».
3. Убедиться, что весь DataMatrix и подпись видны без разворачивания окна.
4. Проверить PNG, PDF и ZPL.

## Commit и отправка в GitHub

Добавлять только файлы текущей задачи:

```powershell
git add <файлы текущей задачи>
git commit -m "краткое описание изменения"
git push origin main
git status --short
git log -1 --oneline --decorate
```

GitHub обновляется только после явного `git push`. Автоматической отправки новых версий нет.

## Не публиковать

В Git не добавлять:

- EXE и установщики Python;
- `build/` и `dist/`;
- `.build*/`, `venv/` и другие виртуальные окружения;
- `tmp/` и `__pycache__/`;
- старые копии исходников;
- снятый с поддержки Win7-комплект;
- секреты и локальные конфигурации.

## Текущая UI-задача

В обычном окне предпросмотр DataMatrix может обрезаться справа. Текущая причина: в `app.py` изображение предпросмотра получает фиксированный предел `430×520`, тогда как правая панель при стартовом размере окна уже.

Предпочтительное исправление:

1. Масштабировать изображение по фактическим размерам `preview_label`.
2. Сохранять пропорции и `Image.Resampling.NEAREST`.
3. Перерисовывать предпросмотр при событии изменения размера окна.
4. Проверить стартовое, уменьшенное и развернутое окно.
