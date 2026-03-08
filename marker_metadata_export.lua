--[[
    marker_metadata_export.lua
    DaVinci Resolve Script — экспорт маркеров + метаданных клипов в CSV,
    с опциональным захватом и экспортом стиллов.

    Поддерживает:
      - Таймлайн-маркеры (timeline markers)
      - Клип-маркеры (clip markers)
      - Опциональный экспорт стиллов (требует Color page)

    ⚠ Для корректной связки стиллов с CSV необходимо в настройках проекта:
      Color → Still Export: включить "Use labels on still export"
      Still label: выбрать "Timeline Timecode"

    Структура экспорта:
      [Выбранная папка]/
      └── [ProjectName]_[TimelineName]_[YYYYMMDD_HHMMSS]/
          ├── markers_[YYYYMMDD_HHMMSS].csv
          └── stills/
              ├── 01.00.28.04_2.1.1.jpg
              └── ...

    Установка:
      macOS:   ~/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Utility/
      Windows: %APPDATA%\Blackmagic Design\DaVinci Resolve\Fusion\Scripts\Utility\

    Запуск: Workspace → Scripts → Utility → marker_metadata_export
--]]

-- ─── Resolve init ────────────────────────────────────────────────────────────

resolve = Resolve()
local projectManager = resolve:GetProjectManager()
local project        = projectManager:GetCurrentProject()
if not project then print("[Error] No project open"); return end

local timeline = project:GetCurrentTimeline()
if not timeline then print("[Error] No timeline open"); return end

-- ─── libavutil (точная конвертация тайм-кодов, по подходу Roger Magnusson) ──

local libavutil = {}

do
    local function load_library(name_pattern)
        local files = bmd.readdir(fu:MapPath("FusionLibs:" .. iif(ffi.os == "Windows", "", "../")) .. name_pattern)
        assert(#files == 1 and files[1].IsDir == false,
            string.format("Couldn't find exact match for pattern \"%s\"", name_pattern))
        return ffi.load(files.Parent .. files[1].Name)
    end

    local lib = load_library(
        iif(ffi.os == "Windows", "avutil*.dll",
            iif(ffi.os == "OSX", "libavutil*.dylib", "libavutil.so")))

    ffi.cdef([[
        typedef struct AVRational { int num; int den; } AVRational;
        typedef struct AVTimecode {
            int start;
            uint32_t flags;
            AVRational rate;
            unsigned fps;
        } AVTimecode;
        int av_timecode_init_from_string(AVTimecode *tc, AVRational rate, const char *str, void *log_ctx);
        char *av_timecode_make_string(const AVTimecode *tc, char *buf, int framenum);
    ]])

    local AV_TIMECODE_FLAG_DROPFRAME     = 0x0001
    local AV_TIMECODE_FLAG_24HOURSMAX    = 0x0002

    local function get_fraction(frame_rate)
        local fr = tonumber(tostring(frame_rate))
        local rates = { 16, 18, 23.976, 24, 25, 29.97, 30, 47.952, 48, 50, 59.94, 60, 72, 95.904, 96, 100, 119.88, 120 }
        for _, r in ipairs(rates) do
            if r == fr or math.floor(r) == fr then
                local is_decimal = r % 1 > 0
                local den = iif(is_decimal, 1001, 100)
                local num = math.ceil(r) * iif(is_decimal, 1000, den)
                return { num = num, den = den }
            end
        end
        return nil
    end

    local function get_decimal(frame_rate)
        local frac = get_fraction(frame_rate)
        return frac and tonumber(string.format("%.3f", frac.num / frac.den)) or nil
    end

    function libavutil.timecode_from_frame(frame, frame_rate, drop_frame)
        local frac = get_fraction(frame_rate)
        if not frac then return tostring(frame) end
        local fps_dec = get_decimal(frame_rate)
        local rate = ffi.new("AVRational", { frac.num, frac.den })
        local tc   = ffi.new("AVTimecode[1]")
        local flags = AV_TIMECODE_FLAG_24HOURSMAX
        if drop_frame == true or drop_frame == 1 or drop_frame == "1" then
            flags = flags + AV_TIMECODE_FLAG_DROPFRAME
        end
        tc[0].start = 0
        tc[0].flags = flags
        tc[0].rate  = rate
        tc[0].fps   = math.ceil(fps_dec)
        local buf = ffi.new("char[64]")
        lib.av_timecode_make_string(tc, buf, frame)
        return ffi.string(buf)
    end
end

-- ─── Стили UI ────────────────────────────────────────────────────────────────

local PRIMARY_COLOR  = "#4C956C"
local HOVER_COLOR    = "#61B15A"
local TEXT_COLOR     = "#ebebeb"
local BORDER_COLOR   = "#3a6ea5"
local SECTION_BG     = "#2A2A2A"
local WARN_COLOR     = "#c0804a"

local PRIMARY_BUTTON = string.format([[
    QPushButton { border: 2px solid %s; border-radius: 8px; background-color: %s; color: #FFF;
        min-height: 35px; font-size: 15px; font-weight: bold; padding: 5px 15px; }
    QPushButton:hover { background-color: %s; border-color: %s; }
    QPushButton:disabled { background-color: #666; border-color: #555; color: #999; }
]], BORDER_COLOR, PRIMARY_COLOR, HOVER_COLOR, PRIMARY_COLOR)

local SECONDARY_BUTTON = string.format([[
    QPushButton { border: 1px solid %s; border-radius: 5px; background-color: %s; color: %s;
        min-height: 28px; font-size: 12px; padding: 3px 10px; }
    QPushButton:hover { background-color: #3A3A3A; }
    QPushButton:disabled { background-color: #555; color: #888; }
]], BORDER_COLOR, SECTION_BG, TEXT_COLOR)

local SECTION = string.format(
    [[ QLabel { color: %s; font-size: 13px; font-weight: bold; padding: 4px 0; } ]], TEXT_COLOR)
local STATUS  = [[ QLabel { color: #a0a0a0; font-size: 11px; padding: 2px 0; } ]]
local WARNING = string.format(
    [[ QLabel { color: %s; font-size: 11px; padding: 2px 0; } ]], WARN_COLOR)

local COMBO = string.format([[
    QComboBox { border: 1px solid %s; border-radius: 4px; padding: 5px;
        background-color: %s; color: %s; min-height: 25px; }
    QComboBox QAbstractItemView { background-color: #1e1e1e; color: %s;
        selection-background-color: #2a3545; }
]], BORDER_COLOR, SECTION_BG, TEXT_COLOR, TEXT_COLOR)

local TREE_STYLE = string.format([[
    QTreeWidget { background-color: #1e1e1e; alternate-background-color: #232323;
        border: 1px solid %s; border-radius: 4px; color: #ebebeb; font-size: 12px; outline: 0; }
    QTreeWidget::item            { height: 24px; padding: 0 4px; }
    QTreeWidget::item:hover      { background: #2a3545; }
    QTreeWidget::item:selected   { background: #2a3545; color: #ebebeb; }
    QHeaderView::section { background: #2A2A2A; color: #aaa; font-size: 11px;
        padding: 3px 6px; border: none; border-bottom: 1px solid %s; }
]], BORDER_COLOR, BORDER_COLOR)

local LINEEDIT_STYLE = string.format([[
    QLineEdit { border: 1px solid %s; border-radius: 4px; padding: 4px 8px;
        background-color: #1e1e1e; color: %s; font-size: 12px; min-height: 24px; }
]], BORDER_COLOR, TEXT_COLOR)

-- ─── Утилиты ─────────────────────────────────────────────────────────────────

local function csv_escape(val)
    if val == nil then return "" end
    local s = tostring(val)
    if s:find('[,"\n\r]') then s = '"' .. s:gsub('"', '""') .. '"' end
    return s
end

local function get_desktop_path()
    local home = os.getenv("HOME") or os.getenv("USERPROFILE") or ""
    return (home ~= "") and (home .. "/Desktop") or "."
end

local function safe_name(s)
    return (s or ""):gsub("[^%w%-%_ ]", "_")
end

local function mkdir(path)
    if ffi.os == "Windows" then
        os.execute('mkdir "' .. path:gsub("/", "\\") .. '" 2>nul')
    else
        os.execute('mkdir -p "' .. path .. '"')
    end
end

-- ─── Параметры таймлайна ──────────────────────────────────────────────────────

local function get_timeline_params()
    local fps_str = timeline:GetSetting("timelineFrameRate")
    local drop    = timeline:GetSetting("timelineDropFrameTimecode")
    local start_f = tonumber(timeline:GetStartFrame()) or 0
    return fps_str, drop, start_f
end

-- ─── Ключи метаданных клипов ──────────────────────────────────────────────────

local function gather_all_clip_meta_keys()
    -- Only include keys that have non-empty data in at least one clip on the timeline
    local keys_with_data = {}
    local tc = timeline:GetTrackCount("video")
    for tr = 1, tc do
        local items = timeline:GetItemListInTrack("video", tr)
        if items then
            for _, item in ipairs(items) do
                local mp = item:GetMediaPoolItem()
                if mp then
                    local props = mp:GetClipProperty() or {}
                    for k, v in pairs(props) do
                        local s = tostring(v or ""):match("^%s*(.-)%s*$")
                        if s ~= "" and s ~= "0" then
                            keys_with_data[k] = true
                        end
                    end
                end
            end
        end
    end
    local keys = {}
    for k in pairs(keys_with_data) do table.insert(keys, k) end
    table.sort(keys)
    return keys
end

local MARKER_FIXED_FIELDS = {
    "Marker Timecode", "Marker Name", "Marker Note", "Marker Color", "Marker Duration",
}

-- ─── Сбор маркеров ───────────────────────────────────────────────────────────

local function find_clip_at_frame(abs_frame)
    local tc = timeline:GetTrackCount("video")
    for tr = 1, tc do
        local items = timeline:GetItemListInTrack("video", tr)
        if items then
            for _, item in ipairs(items) do
                local s = item:GetStart()
                local e = s + item:GetDuration() - 1
                if abs_frame >= s and abs_frame <= e then
                    local mp = item:GetMediaPoolItem()
                    if mp then return mp:GetName() or "N/A", mp:GetClipProperty() or {} end
                    return "N/A", {}
                end
            end
        end
    end
    return "N/A", {}
end

local function make_row(abs_frame, marker, clip_name, clip_props, fps_str, drop)
    return {
        ["Marker Timecode"] = libavutil.timecode_from_frame(abs_frame, fps_str, drop),
        ["Marker Name"]     = marker.name     or "",
        ["Marker Note"]     = marker.note     or "",
        ["Marker Color"]    = marker.color    or "",
        ["Marker Duration"] = marker.duration or "",
        ["Clip Name"]       = clip_name,
        clip_props          = clip_props,
        abs_frame           = abs_frame,
    }
end

local function collect_timeline_markers(fps_str, drop, start_frame)
    local rows   = {}
    local mdict  = timeline:GetMarkers() or {}
    local sorted = {}
    for fid, m in pairs(mdict) do
        table.insert(sorted, { frame_id = tonumber(fid), marker = m })
    end
    table.sort(sorted, function(a, b) return a.frame_id < b.frame_id end)
    for _, e in ipairs(sorted) do
        local abs = start_frame + e.frame_id
        local cn, cp = find_clip_at_frame(abs)
        table.insert(rows, make_row(abs, e.marker, cn, cp, fps_str, drop))
    end
    return rows
end

local function collect_clip_markers(fps_str, drop, start_frame)
    local rows     = {}
    local tc_count = timeline:GetTrackCount("video")
    for tr = 1, tc_count do
        local items = timeline:GetItemListInTrack("video", tr)
        if items then
            for _, item in ipairs(items) do
                local cm = item:GetMarkers()
                if cm and next(cm) ~= nil then
                    local mp = item:GetMediaPoolItem()
                    local cn  = mp and mp:GetName() or "N/A"
                    local cp  = (mp and mp:GetClipProperty()) or {}
                    local cs  = item:GetStart()
                    for offset, m in pairs(cm) do
                        local abs = cs + tonumber(offset)
                        table.insert(rows, make_row(abs, m, cn, cp, fps_str, drop))
                    end
                end
            end
        end
    end
    table.sort(rows, function(a, b) return a.abs_frame < b.abs_frame end)
    return rows
end

-- ─── Экспорт стиллов ─────────────────────────────────────────────────────────

local function export_stills(rows, fps_str, drop, stills_dir, format)
    local gallery    = project:GetGallery()
    local orig_album = gallery:GetCurrentStillAlbum()

    -- Создаём изолированный альбом для этой сессии
    local album_name   = "Marker Export " .. os.date("%Y%m%d_%H%M%S")
    local export_album = gallery:CreateGalleryStillAlbum(album_name)
    gallery:SetCurrentStillAlbum(export_album)

    -- Сохраняем текущую страницу и тайм-код
    local prev_page = resolve:GetCurrentPage()
    local prev_tc   = timeline:GetCurrentTimecode()
    resolve:OpenPage("color")

    local grabbed = {}
    for _, row in ipairs(rows) do
        local tc = libavutil.timecode_from_frame(row.abs_frame, fps_str, drop)
        if tc and timeline:SetCurrentTimecode(tc) then
            local still = timeline:GrabStill()
            if still then
                table.insert(grabbed, still)
            else
                print("[WARN] GrabStill failed at " .. tc)
            end
        else
            print("[WARN] SetCurrentTimecode failed for frame " .. tostring(row.abs_frame))
        end
    end

    if #grabbed > 0 then
        -- Workaround: reselect album перед ExportStills (баг Resolve с Timelines album)
        for _, a in ipairs(gallery:GetGalleryStillAlbums()) do
            if a ~= export_album then gallery:SetCurrentStillAlbum(a); break end
        end
        gallery:SetCurrentStillAlbum(export_album)

        -- Пустой prefix → Resolve использует лейблы из настроек проекта (Timeline Timecode)
        assert(
            export_album:ExportStills(grabbed, stills_dir, "", format),
            "ExportStills failed"
        )
    end

    -- Восстанавливаем состояние
    resolve:OpenPage(prev_page)
    if prev_tc then timeline:SetCurrentTimecode(prev_tc) end
    gallery:SetCurrentStillAlbum(orig_album)

    print(string.format("[INFO] %d stills grabbed → album '%s'", #grabbed, album_name))
end

-- ─── Сопоставление стиллов с CSV ─────────────────────────────────────────────

-- "01:00:28:04" → "01.00.28.04"
local function tc_to_prefix(tc)
    return tc:gsub("[:;]", ".")
end

-- Сканируем папку стиллов, строим map: prefix → filename
local function build_stills_map(stills_dir)
    local map   = {}
    local files = bmd.readdir(stills_dir .. "/*")
    if files then
        for _, f in ipairs(files) do
            if not f.IsDir and not f.Name:lower():match("%.drx$") then
                -- Имя файла: "01.00.28.04_2.1.1.jpg" → prefix = "01.00.28.04"
                local prefix = f.Name:match("^([%d%.]+)_")
                if prefix then map[prefix] = f.Name end
            end
        end
    end
    return map
end

-- ─── Запись CSV ──────────────────────────────────────────────────────────────

local function write_csv(rows, filtered_fixed, include_clip_name, meta_fields,
                          csv_path, stills_dir, with_stills)
    local stills_map = {}
    if with_stills and stills_dir then
        stills_map = build_stills_map(stills_dir)
    end

    local file, err = io.open(csv_path, "w")
    if not file then return false, tostring(err) end

    file:write("\xEF\xBB\xBF")  -- BOM для Excel

    -- Заголовки
    local headers = {}
    for _, f in ipairs(filtered_fixed) do table.insert(headers, f) end
    if include_clip_name             then table.insert(headers, "Clip Name") end
    for _, k in ipairs(meta_fields)  do table.insert(headers, k) end
    if with_stills then
        table.insert(headers, "Still Filename")
        table.insert(headers, "Still Path")
    end

    local h_line = {}
    for _, h in ipairs(headers) do table.insert(h_line, csv_escape(h)) end
    file:write(table.concat(h_line, ",") .. "\n")

    -- Строки данных
    for _, row in ipairs(rows) do
        local line = {}
        for _, f in ipairs(filtered_fixed) do
            table.insert(line, csv_escape(row[f]))
        end
        if include_clip_name then
            table.insert(line, csv_escape(row["Clip Name"]))
        end
        for _, k in ipairs(meta_fields) do
            table.insert(line, csv_escape(row.clip_props and row.clip_props[k] or ""))
        end
        if with_stills then
            local prefix   = tc_to_prefix(row["Marker Timecode"] or "")
            local filename = stills_map[prefix] or ""
            local fullpath = (filename ~= "" and stills_dir) and (stills_dir .. "/" .. filename) or ""
            table.insert(line, csv_escape(filename))
            table.insert(line, csv_escape(fullpath))
        end
        file:write(table.concat(line, ",") .. "\n")
    end

    file:close()
    return true, nil
end

-- ─── Marker colors (filter) ───────────────────────────────────────────────────

local function get_used_marker_colors(marker_type)
    local colors_set = {}
    if marker_type == "Timeline Markers" then
        for _, m in pairs(timeline:GetMarkers() or {}) do
            if m.color and m.color ~= "" then colors_set[m.color] = true end
        end
    else
        local tc = timeline:GetTrackCount("video")
        for tr = 1, tc do
            local items = timeline:GetItemListInTrack("video", tr)
            if items then
                for _, item in ipairs(items) do
                    for _, m in pairs(item:GetMarkers() or {}) do
                        if m.color and m.color ~= "" then colors_set[m.color] = true end
                    end
                end
            end
        end
    end
    local colors = {}
    for c in pairs(colors_set) do table.insert(colors, c) end
    table.sort(colors)
    return colors
end

local function count_markers(marker_type, color_filter)
    local count = 0
    local function matches(m)
        return color_filter == "All Colors" or m.color == color_filter
    end
    if marker_type == "Timeline Markers" then
        for _, m in pairs(timeline:GetMarkers() or {}) do
            if matches(m) then count = count + 1 end
        end
    else
        local tc = timeline:GetTrackCount("video")
        for tr = 1, tc do
            local items = timeline:GetItemListInTrack("video", tr)
            if items then
                for _, item in ipairs(items) do
                    for _, m in pairs(item:GetMarkers() or {}) do
                        if matches(m) then count = count + 1 end
                    end
                end
            end
        end
    end
    return count
end

-- ─── Форматы стиллов ─────────────────────────────────────────────────────────

local STILL_FORMATS = { "jpg", "png", "tif", "dpx", "cin", "bmp" }

-- ─── UI ──────────────────────────────────────────────────────────────────────

local ui   = fu.UIManager
local disp = bmd.UIDispatcher(ui)

local all_meta_keys = gather_all_clip_meta_keys()

local win = disp:AddWindow({
    ID          = "MarkerExportWin",
    WindowTitle = "Marker Metadata Export",
    Geometry    = { 300, 120, 540, 740 },
    Spacing     = 8,

    ui:VGroup{
        ui:Label{
            Weight = 0,
            Text   = "Marker Metadata Export",
            StyleSheet = string.format(
                [[ QLabel { color: %s; font-size: 15px; font-weight: bold; padding: 6px 0 2px 0; } ]],
                TEXT_COLOR)
        },

        -- Marker source and color filter
        ui:Label{ Weight = 0, Text = "Marker Source", StyleSheet = SECTION },
        ui:ComboBox{ ID = "MarkerType", Weight = 0, StyleSheet = COMBO },
        ui:HGroup{
            Weight = 0,
            ui:Label{ Weight = 0, Text = "Color filter:", StyleSheet = STATUS },
            ui:HGap(6),
            ui:ComboBox{ ID = "ColorFilter", Weight = 1, StyleSheet = COMBO },
        },
        ui:Label{
            ID         = "MarkerCount",
            Weight     = 0,
            Text       = "",
            Alignment  = { AlignLeft = true },
            StyleSheet = STATUS,
        },

        ui:VGap(4),

        -- Таблица полей (растягивается при изменении размера окна)
        ui:Label{ Weight = 0, Text = "Fields to export", StyleSheet = SECTION },
        ui:HGroup{
            Weight = 0,
            ui:Button{ ID = "SelectAll",   Text = "Select All",   MinimumSize = {90, 26}, StyleSheet = SECONDARY_BUTTON },
            ui:Button{ ID = "DeselectAll", Text = "Deselect All", MinimumSize = {90, 26}, StyleSheet = SECONDARY_BUTTON },
        },
        ui:Tree{
            ID                   = "FieldTree",
            Weight               = 1,
            AlternatingRowColors = true,
            RootIsDecorated      = true,
            StyleSheet           = TREE_STYLE,
            HeaderHidden         = false,
        },

        ui:VGap(4),

        -- Опция экспорта стиллов
        ui:Label{ Weight = 0, Text = "Stills Export (optional)", StyleSheet = SECTION },
        ui:HGroup{
            Weight = 0,
            ui:CheckBox{
                ID      = "StillsCheck",
                Weight  = 0,
                Text    = "Export stills at markers",
                Checked = false,
            },
            ui:HGap(10),
            ui:Label{ Weight = 0, Text = "Format:", StyleSheet = STATUS },
            ui:ComboBox{ ID = "StillFormat", Weight = 0, StyleSheet = COMBO, Enabled = false },
        },
        ui:Label{
            ID         = "StillsWarning",
            Weight     = 0,
            Hidden     = true,
            WordWrap   = true,
            Text       = "⚠  Required settings:\n" ..
                         "Project Settings → General Options → Color:\n" ..
                         "  • Automatically label gallery stills using: Timeline Timecode\n" ..
                         "  • Append still number on export: As Suffix\n" ..
                         "Color page → Stills export options:\n" ..
                         "  • Use labels on still export",
            StyleSheet = string.format([[
                QLabel {
                    color: %s;
                    font-size: 11px;
                    padding: 6px 8px;
                    border: 1px solid %s;
                    border-radius: 4px;
                    background-color: #2a1f10;
                    line-height: 1.5;
                }
            ]], WARN_COLOR, WARN_COLOR),
        },

        ui:VGap(4),

        -- Папка для сохранения
        ui:Label{ Weight = 0, Text = "Output folder", StyleSheet = SECTION },
        ui:HGroup{
            Weight = 0,
            ui:LineEdit{
                ID         = "OutputPath",
                Text       = get_desktop_path(),
                StyleSheet = LINEEDIT_STYLE,
            },
            ui:Button{
                ID          = "BrowseBtn",
                Text        = "Browse…",
                MinimumSize = {70, 28},
                MaximumSize = {70, 28},
                StyleSheet  = SECONDARY_BUTTON,
            },
        },

        ui:VGap(4),

        -- Статус
        ui:Label{ ID = "Status", Weight = 0, Text = "",
            Alignment = { AlignCenter = true }, StyleSheet = STATUS },

        -- Кнопки
        ui:HGroup{
            Weight = 0,
            ui:Button{ ID = "Export", Text = "Export", StyleSheet = PRIMARY_BUTTON },
            ui:Button{ ID = "Cancel", Text = "Cancel", MinimumSize = {90, 35}, StyleSheet = SECONDARY_BUTTON },
        },
    }
})

local itm = win:GetItems()

-- ─── Заполнение контролов ─────────────────────────────────────────────────────

itm.MarkerType:AddItem("Clip Markers")
itm.MarkerType:AddItem("Timeline Markers")

for _, fmt in ipairs(STILL_FORMATS) do itm.StillFormat:AddItem(fmt) end

-- Populate ColorFilter and update marker count
local function refresh_color_filter()
    local marker_type = itm.MarkerType.CurrentText
    local prev_color  = itm.ColorFilter.CurrentText

    itm.ColorFilter:Clear()
    itm.ColorFilter:AddItem("All Colors")
    local colors = get_used_marker_colors(marker_type)
    for _, c in ipairs(colors) do itm.ColorFilter:AddItem(c) end

    itm.ColorFilter.CurrentText = prev_color
    if itm.ColorFilter.CurrentText ~= prev_color then
        itm.ColorFilter.CurrentIndex = 0
    end
end

local function update_marker_count()
    local marker_type  = itm.MarkerType.CurrentText
    local color_filter = itm.ColorFilter.CurrentText or "All Colors"
    local count = count_markers(marker_type, color_filter)
    if count == 0 then
        itm.MarkerCount.Text = "No markers found"
    elseif color_filter == "All Colors" then
        itm.MarkerCount.Text = string.format("Found: %d marker%s", count, count == 1 and "" or "s")
    else
        itm.MarkerCount.Text = string.format("Found: %d %s marker%s", count, color_filter, count == 1 and "" or "s")
    end
end

refresh_color_filter()
update_marker_count()

-- ─── Дерево полей ─────────────────────────────────────────────────────────────

itm.FieldTree:SetColumnCount(1)
itm.FieldTree:SetHeaderLabels({ "Field" })

local marker_group = itm.FieldTree:NewItem()
marker_group.Text[0]       = "Marker Fields"
marker_group.CheckState[0] = "Checked"
marker_group.Flags         = { ItemIsUserCheckable = true, ItemIsEnabled = true }
itm.FieldTree:AddTopLevelItem(marker_group)
for _, field in ipairs(MARKER_FIXED_FIELDS) do
    local child = itm.FieldTree:NewItem()
    child.Text[0]       = field
    child.CheckState[0] = "Checked"
    child.Flags         = { ItemIsUserCheckable = true, ItemIsEnabled = true }
    marker_group:AddChild(child)
end
marker_group.Expanded = true

local clipname_group = itm.FieldTree:NewItem()
clipname_group.Text[0]       = "Clip Info"
clipname_group.CheckState[0] = "Checked"
clipname_group.Flags         = { ItemIsUserCheckable = true, ItemIsEnabled = true }
itm.FieldTree:AddTopLevelItem(clipname_group)
local clipname_child = itm.FieldTree:NewItem()
clipname_child.Text[0]       = "Clip Name"
clipname_child.CheckState[0] = "Checked"
clipname_child.Flags         = { ItemIsUserCheckable = true, ItemIsEnabled = true }
clipname_group:AddChild(clipname_child)
clipname_group.Expanded = true

local meta_group = itm.FieldTree:NewItem()
meta_group.Text[0]       = "Clip Metadata"
meta_group.CheckState[0] = "Checked"
meta_group.Flags         = { ItemIsUserCheckable = true, ItemIsEnabled = true }
itm.FieldTree:AddTopLevelItem(meta_group)
for _, key in ipairs(all_meta_keys) do
    local child = itm.FieldTree:NewItem()
    child.Text[0]       = key
    child.CheckState[0] = "Checked"
    child.Flags         = { ItemIsUserCheckable = true, ItemIsEnabled = true }
    meta_group:AddChild(child)
end
meta_group.Expanded = false

-- ─── Хелперы дерева ──────────────────────────────────────────────────────────

local function set_children_check(parent, state)
    for i = 0, parent:ChildCount() - 1 do
        local child = parent:Child(i)
        child.CheckState[0] = state
        set_children_check(child, state)
    end
end

local function get_selected_fields()
    local marker_fields, meta_fields = {}, {}
    local include_clip_name = false
    for i = 0, marker_group:ChildCount() - 1 do
        local c = marker_group:Child(i)
        if c.CheckState[0] == "Checked" then table.insert(marker_fields, c.Text[0]) end
    end
    if clipname_child.CheckState[0] == "Checked" then include_clip_name = true end
    for i = 0, meta_group:ChildCount() - 1 do
        local c = meta_group:Child(i)
        if c.CheckState[0] == "Checked" then table.insert(meta_fields, c.Text[0]) end
    end
    return marker_fields, include_clip_name, meta_fields
end

-- Сохраняем оригинальный порядок MARKER_FIXED_FIELDS
local function filter_fixed_fields(marker_fields)
    local set = {}
    for _, f in ipairs(marker_fields) do set[f] = true end
    local out = {}
    for _, f in ipairs(MARKER_FIXED_FIELDS) do
        if set[f] then table.insert(out, f) end
    end
    return out
end

-- ─── Обработчики событий ─────────────────────────────────────────────────────

win.On.MarkerType.CurrentIndexChanged = function()
    refresh_color_filter()
    update_marker_count()
end

win.On.ColorFilter.CurrentIndexChanged = function()
    update_marker_count()
end

win.On.BrowseBtn.Clicked = function()
    local selected = tostring(fu:RequestDir(itm.OutputPath.Text or get_desktop_path()))
    if selected and selected ~= "" and selected ~= "nil" then
        itm.OutputPath.Text = selected
    end
end

win.On.SelectAll.Clicked = function()
    for _, g in ipairs({ marker_group, clipname_group, meta_group }) do
        g.CheckState[0] = "Checked"; set_children_check(g, "Checked")
    end
end

win.On.DeselectAll.Clicked = function()
    for _, g in ipairs({ marker_group, clipname_group, meta_group }) do
        g.CheckState[0] = "Unchecked"; set_children_check(g, "Unchecked")
    end
end

win.On.FieldTree.ItemClicked = function(ev)
    set_children_check(ev.item, ev.item.CheckState[0])
end

win.On.StillsCheck.Clicked = function()
    local checked = itm.StillsCheck.Checked
    itm.StillFormat.Enabled  = checked
    itm.StillsWarning.Hidden = not checked

    -- Расширяем/сжимаем окно по высоте чтобы вместить предупреждение
    local g = win.Geometry
    local delta = 110  -- примерная высота блока предупреждения
    if checked then
        win.Geometry = { g[1], g[2], g[3], g[4] + delta }
    else
        win.Geometry = { g[1], g[2], g[3], math.max(g[4] - delta, 400) }
    end
    win:RecalcLayout()
    win:Update()
end

win.On.Export.Clicked = function()
    local marker_fields, include_clip_name, meta_fields = get_selected_fields()
    if #marker_fields == 0 and not include_clip_name and #meta_fields == 0 then
        itm.Status.Text = "⚠ Select at least one field."
        return
    end

    local base_dir = (itm.OutputPath.Text or get_desktop_path()):gsub("[/\\]+$", "")
    local fps_str, drop, start_frame = get_timeline_params()
    local timestamp   = os.date("%Y%m%d_%H%M%S")
    local folder_name = safe_name(project:GetName()) .. "_" .. safe_name(timeline:GetName()) .. "_" .. timestamp
    local export_dir  = base_dir .. "/" .. folder_name
    local stills_dir  = export_dir .. "/stills"

    mkdir(export_dir)

    -- Собираем маркеры
    local rows
    if itm.MarkerType.CurrentText == "Timeline Markers" then
        rows = collect_timeline_markers(fps_str, drop, start_frame)
    else
        rows = collect_clip_markers(fps_str, drop, start_frame)
    end

    -- Apply color filter
    local color_filter = itm.ColorFilter.CurrentText or "All Colors"
    if color_filter ~= "All Colors" then
        local filtered = {}
        for _, row in ipairs(rows) do
            if row["Marker Color"] == color_filter then
                table.insert(filtered, row)
            end
        end
        rows = filtered
    end

    if #rows == 0 then
        itm.Status.Text = "⚠ No markers found for selected filter."
        return
    end

    -- Экспорт стиллов (до CSV, чтобы сопоставление было актуальным)
    local with_stills = itm.StillsCheck.Checked
    if with_stills then
        mkdir(stills_dir)
        itm.Status.Text = "Grabbing stills… please wait"
        local ok, err = pcall(export_stills, rows, fps_str, drop, stills_dir,
            STILL_FORMATS[itm.StillFormat.CurrentIndex + 1] or "jpg")
        if not ok then
            print("[ERROR] Stills export failed: " .. tostring(err))
            itm.Status.Text = "⚠ Stills failed: " .. tostring(err) .. " — CSV will still be saved."
            with_stills = false  -- CSV пишем без колонок стиллов
        end
    end

    -- Записываем CSV
    local csv_path = export_dir .. "/markers_" .. timestamp .. ".csv"
    local ok, err  = write_csv(
        rows, filter_fixed_fields(marker_fields), include_clip_name, meta_fields,
        csv_path, stills_dir, with_stills
    )

    if not ok then
        itm.Status.Text = "⚠ CSV error: " .. tostring(err)
        return
    end

    itm.Status.Text = string.format("✓ Done! %d rows → %s", #rows, export_dir)
    print("[DONE] Export complete: " .. export_dir)
end

win.On.Cancel.Clicked        = function() disp:ExitLoop() end
win.On.MarkerExportWin.Close = function() disp:ExitLoop() end

-- ─── Запуск ───────────────────────────────────────────────────────────────────

win:Show()
disp:RunLoop()
win:Hide()
