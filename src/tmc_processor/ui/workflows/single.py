"""Single-file workflow stage implementations."""

from __future__ import annotations

from tmc_processor.ui.workflow_context import WorkflowContext

import streamlit as st


def render_single_data(*, context: WorkflowContext) -> None:
    """Render the Single Data stage."""
    _render_action_hint = context.operations._render_action_hint
    _render_section_header = context.operations._render_section_header
    direction_cols = context.direction_cols
    info_cols = context.info_cols
    report_cols = context.report_cols
    road_cols = context.road_cols
    setup_left = context.setup_left
    setup_right = context.setup_right
    uploaded_file = context.uploaded_file

    _render_section_header("Data", "Upload the source workbook and enter project/report setup.")
    if uploaded_file is None:
        _render_action_hint("Start by uploading a TMC Excel workbook above.")

    setup_left, setup_right = st.columns([1.15, 1])
    with setup_left:
        with st.container(border=True):
            _render_section_header("ข้อมูลโครงการและรายงาน", "ข้อมูลหลักสำหรับปกและหัวรายงาน")
            report_cols = st.columns(2)
            project_name = report_cols[0].text_input("ชื่อโครงการ", key="project_name_input")
            tmc_id = report_cols[1].text_input("TMC ID", key="tmc_id_input")
            tmc_title = st.text_input("ชื่อจุดนับ", key="tmc_title_input")
            info_cols = st.columns(2)
            survey_point = info_cols[0].text_input("จุดสำรวจ", key="survey_point_input")
            survey_date_text = info_cols[1].text_input("วันที่สำรวจ", key="survey_date_text_input")
            weather = info_cols[0].text_input("สภาพอากาศ", key="weather_input")
            responsible_party = info_cols[1].text_input("ผู้รับผิดชอบ", key="responsible_party_input")

    with setup_right:
        with st.container(border=True):
            _render_section_header("ป้ายปลายทางและถนน", "ข้อความที่ใช้ใน Diagram และรายงาน")
            st.caption("ป้ายปลายทาง")
            direction_cols = st.columns(2)
            north_label = direction_cols[0].text_input("ป้ายปลายทางด้านเหนือ", key="north_label_input")
            south_label = direction_cols[1].text_input("ป้ายปลายทางด้านใต้", key="south_label_input")
            east_label = direction_cols[0].text_input("ป้ายปลายทางด้านตะวันออก", key="east_label_input")
            west_label = direction_cols[1].text_input("ป้ายปลายทางด้านตะวันตก", key="west_label_input")
            st.caption("ชื่อถนน / ทางหลวง")
            road_cols = st.columns(2)
            north_road = road_cols[0].text_input("ชื่อถนนด้านเหนือ", key="north_road_input")
            south_road = road_cols[1].text_input("ชื่อถนนด้านใต้", key="south_road_input")
            east_road = road_cols[0].text_input("ชื่อถนนด้านตะวันออก", key="east_road_input")
            west_road = road_cols[1].text_input("ชื่อถนนด้านตะวันตก", key="west_road_input")
            caption_text = st.text_input("คำบรรยายรูป Diagram", key="caption_text_input")
            show_u_turn = st.checkbox("แสดง movement กลับรถ", key="show_u_turn_checkbox")
def render_single_mapping(*, context: WorkflowContext) -> None:
    """Render the Single Mapping stage."""
    BytesIO = context.operations.BytesIO
    EXCEL_MIME = context.operations.EXCEL_MIME
    MAPPING_PRESET_MIME = context.operations.MAPPING_PRESET_MIME
    MAPPING_SCHEME_LOCK_CAPTION = context.operations.MAPPING_SCHEME_LOCK_CAPTION
    MAPPING_SOURCE_DEFAULT_PREVIEW = context.operations.MAPPING_SOURCE_DEFAULT_PREVIEW
    MAPPING_SOURCE_MAPPING_EXCEL = context.operations.MAPPING_SOURCE_MAPPING_EXCEL
    MAPPING_SOURCE_MAPPING_PRESET = context.operations.MAPPING_SOURCE_MAPPING_PRESET
    MAPPING_SOURCE_USER_EDITOR = context.operations.MAPPING_SOURCE_USER_EDITOR
    MOVEMENT_SCHEMES = context.operations.MOVEMENT_SCHEMES
    MappingPresetError = context.operations.MappingPresetError
    _apply_mapping_editor_widget_state = context.operations._apply_mapping_editor_widget_state
    _basic_mapping_widget_state_exists = context.operations._basic_mapping_widget_state_exists
    _clear_mapping_for_scheme_change = context.operations._clear_mapping_for_scheme_change
    _current_mapping_scheme = context.operations._current_mapping_scheme
    _is_v2_scheme = context.operations._is_v2_scheme
    _mapping_aggregation_preview = context.operations._mapping_aggregation_preview
    _mapping_editor_changed = context.operations._mapping_editor_changed
    _mapping_editor_column_config = context.operations._mapping_editor_column_config
    _mapping_editor_frame = context.operations._mapping_editor_frame
    _mapping_editor_key = context.operations._mapping_editor_key
    _mapping_from_editor_widget_state = context.operations._mapping_from_editor_widget_state
    _mapping_issue_display = context.operations._mapping_issue_display
    _mapping_rows_are_committed = context.operations._mapping_rows_are_committed
    _mapping_source = context.operations._mapping_source
    _mapping_workspace_counts = context.operations._mapping_workspace_counts
    _merge_mapping_editor_result = context.operations._merge_mapping_editor_result
    _movement_code_options_for_scheme = context.operations._movement_code_options_for_scheme
    _render_action_hint = context.operations._render_action_hint
    _render_alert = context.operations._render_alert
    _render_basic_mapping_editor = context.operations._render_basic_mapping_editor
    _render_download_button = context.operations._render_download_button
    _render_empty_state = context.operations._render_empty_state
    _render_mapping_scheme_status = context.operations._render_mapping_scheme_status
    _render_metric_strip = context.operations._render_metric_strip
    _render_section_header = context.operations._render_section_header
    _scheme_select_label = context.operations._scheme_select_label
    _set_current_mapping_scheme = context.operations._set_current_mapping_scheme
    _set_mapping_source = context.operations._set_mapping_source
    _single_file_mapping_issues = context.operations._single_file_mapping_issues
    _single_file_processing_block_reason = context.operations._single_file_processing_block_reason
    _sync_single_workflow_from_state = context.operations._sync_single_workflow_from_state
    _thai_aggregation_message = context.operations._thai_aggregation_message
    _thai_mapping_control_warning = context.operations._thai_mapping_control_warning
    aggregation_message = context.aggregation_message
    aggregation_preview = context.aggregation_preview
    apply_mapping_preset_to_detected_sheets = context.operations.apply_mapping_preset_to_detected_sheets
    apply_result = context.apply_result
    apply_saved_mapping_to_sheets = context.operations.apply_saved_mapping_to_sheets
    build_mapping_preset = context.operations.build_mapping_preset
    default_mapping = context.default_mapping
    default_mapping_for_sheets = context.operations.default_mapping_for_sheets
    detect_mapping_preset_scheme = context.operations.detect_mapping_preset_scheme
    detected_scheme = context.detected_scheme
    detected_sheet_names = context.detected_sheet_names
    download_buffer = context.operations.download_buffer
    edited_frame = context.edited_frame
    editor_changed = context.editor_changed
    editor_frame = context.editor_frame
    editor_key = context.editor_key
    exc = context.exc
    export_mode = context.export_mode
    file_bytes = context.file_bytes
    hashlib = context.operations.hashlib
    issue = context.issue
    load_mapping_preset = context.operations.load_mapping_preset
    loaded_excel = context.loaded_excel
    loaded_preset = context.loaded_preset
    mapping = context.mapping
    mapping_control_warnings = context.operations.mapping_control_warnings
    mapping_counts = context.mapping_counts
    mapping_editor_version = context.mapping_editor_version
    mapping_excel_col = context.mapping_excel_col
    mapping_issues = context.mapping_issues
    mapping_preset_col = context.mapping_preset_col
    mapping_preset_upload = context.mapping_preset_upload
    mapping_rows_committed = context.mapping_rows_committed
    mapping_scheme = context.mapping_scheme
    mapping_to_excel_bytes = context.operations.mapping_to_excel_bytes
    mapping_upload = context.mapping_upload
    mapping_upload_bytes = context.mapping_upload_bytes
    mapping_upload_identity = context.mapping_upload_identity
    mapping_view = context.mapping_view
    movement_aggregation_messages = context.operations.movement_aggregation_messages
    movement_code_options = context.movement_code_options
    normalize_approach_movement_mapping = context.operations.normalize_approach_movement_mapping
    pd = context.operations.pd
    preset_bytes = context.preset_bytes
    preset_info = context.preset_info
    preset_name_seed = context.preset_name_seed
    preset_source = context.preset_source
    preset_upload_bytes = context.preset_upload_bytes
    preset_upload_identity = context.preset_upload_identity
    preview_summary = context.preview_summary
    process_block_reason = context.process_block_reason
    read_mapping_excel_with_metadata = context.operations.read_mapping_excel_with_metadata
    safe_mapping_preset_filename = context.operations.safe_mapping_preset_filename
    scheme_validation_issues = context.scheme_validation_issues
    selected_scheme = context.selected_scheme
    serialize_mapping_preset = context.operations.serialize_mapping_preset
    uploaded_file = context.uploaded_file
    use_basic_controls = context.use_basic_controls
    validate_mapping_scheme = context.operations.validate_mapping_scheme
    warning_message = context.warning_message

    _render_section_header(
        "Mapping",
        "Map detected source sheets to movement directions. Analysis is performed in Analyze.",
    )
    if uploaded_file is None:
        _render_empty_state(
            "ยังไม่มีไฟล์สำรวจ",
            "Start from the Data stage and upload a TMC Excel workbook.",
        )
    elif not detected_sheet_names:
        _render_alert('ไม่พบ Sheet ทิศทางจากไฟล์สำรวจ ควรมีชื่อ Sheet เช่น "ทิศ 1", "ทิศ 2", หรือ "ทิศ 2+3"', "warning")
    else:
        default_mapping = default_mapping_for_sheets(detected_sheet_names)
        if st.session_state.get("mapping_table") is not None:
            default_mapping = apply_saved_mapping_to_sheets(detected_sheet_names, pd.DataFrame(st.session_state["mapping_table"]))
        mapping_scheme = _current_mapping_scheme()
        mapping_rows_committed = _mapping_rows_are_committed()
        with st.expander("Advanced mapping controls", expanded=False):
            selected_scheme = st.selectbox(
                "ระบบรหัส Movement",
                options=MOVEMENT_SCHEMES,
                index=MOVEMENT_SCHEMES.index(mapping_scheme),
                format_func=_scheme_select_label,
                disabled=mapping_rows_committed,
                key="movement_code_scheme_selector",
                help="เลือกได้ก่อนโหลดหรือกรอก Mapping; ถ้ามี Mapping แล้วให้ล้างหรือโหลด Mapping ใหม่เพื่อไม่ตีความรหัสเดิมผิด",
            )
            if selected_scheme != mapping_scheme and not mapping_rows_committed:
                _set_current_mapping_scheme(selected_scheme)
                mapping_scheme = selected_scheme
                st.session_state["mapping_editor_version"] = int(st.session_state.get("mapping_editor_version", 0) or 0) + 1
            if mapping_rows_committed:
                st.caption(f"Detected movement_code_scheme: {mapping_scheme}")
                st.caption(MAPPING_SCHEME_LOCK_CAPTION)
                if st.button("ล้าง Mapping เพื่อเปลี่ยนระบบรหัส", key="clear_mapping_for_scheme_change"):
                    _clear_mapping_for_scheme_change()
                    st.rerun()
            else:
                st.caption(f"Selected movement_code_scheme: {mapping_scheme}")
        preset_name_seed = st.session_state.get("tmc_id_input") or st.session_state.get("tmc_title_input") or uploaded_file.name
        preset_source = pd.DataFrame(st.session_state.get("mapping_table") or default_mapping.to_dict("records"))
        preset_bytes = serialize_mapping_preset(
            build_mapping_preset(
                preset_source,
                preset_name=str(preset_name_seed or "TMC Mapping Preset"),
                movement_code_scheme=mapping_scheme,
            )
        )
        st.session_state["tmc_mapping_preset_bytes"] = preset_bytes
        st.session_state["tmc_mapping_preset_filename"] = safe_mapping_preset_filename(preset_name_seed)

        with st.container(border=True):
            _render_section_header("นำเข้า/ส่งออก Mapping", "เลือกใช้ Mapping Excel สำหรับแก้ไขใน Excel หรือ Mapping Preset สำหรับนำค่าที่ตั้งไว้กลับมาใช้ซ้ำ")
            mapping_excel_col, mapping_preset_col = st.columns(2)
            with mapping_excel_col:
                st.markdown("**Mapping Excel**")
                st.caption("สำหรับกรอกหรือแก้ไข Mapping ด้วย Excel")
                mapping_upload = st.file_uploader(
                    "โหลดไฟล์ Mapping Excel",
                    type=["xlsx", "xlsm", "xls"],
                    key="mapping_upload",
                )
                _render_download_button(
                    "ดาวน์โหลดเทมเพลต Mapping",
                    mapping_to_excel_bytes(default_mapping, movement_code_scheme=mapping_scheme),
                    "tmc_mapping_template.xlsx",
                    EXCEL_MIME,
                )
            with mapping_preset_col:
                st.markdown("**Mapping Preset**")
                st.caption("สำหรับบันทึก Mapping ที่ตั้งค่าแล้วและนำกลับมาใช้ซ้ำในโปรแกรม")
                mapping_preset_upload = st.file_uploader(
                    "เปิด Mapping Preset",
                    type=["json"],
                    key="mapping_preset_upload",
                )
                st.download_button(
                    "ดาวน์โหลด Mapping Preset",
                    data=download_buffer(preset_bytes),
                    file_name=st.session_state["tmc_mapping_preset_filename"],
                    mime=MAPPING_PRESET_MIME,
                    key="download_mapping_preset",
                )
        if mapping_upload is not None:
            try:
                mapping_upload_bytes = mapping_upload.getvalue()
                mapping_upload_identity = (mapping_upload.name, hashlib.sha256(mapping_upload_bytes).hexdigest())
                if (
                    st.session_state.get("tmc_mapping_upload_identity") != mapping_upload_identity
                    and st.session_state.get("tmc_mapping_upload_ignored_identity") != mapping_upload_identity
                ):
                    st.session_state.pop("tmc_mapping_upload_ignored_identity", None)
                    loaded_excel = read_mapping_excel_with_metadata(BytesIO(mapping_upload_bytes))
                    default_mapping = apply_saved_mapping_to_sheets(detected_sheet_names, loaded_excel.mapping)
                    _set_current_mapping_scheme(loaded_excel.movement_code_scheme)
                    st.session_state["tmc_loaded_mapping_code_scheme"] = loaded_excel.movement_code_scheme
                    st.session_state["mapping_table"] = default_mapping.to_dict("records")
                    st.session_state["mapping_editor_version"] = int(st.session_state.get("mapping_editor_version", 0) or 0) + 1
                    st.session_state["tmc_mapping_table_from_session"] = False
                    _set_mapping_source(MAPPING_SOURCE_MAPPING_EXCEL)
                    st.session_state["tmc_mapping_upload_identity"] = mapping_upload_identity
                    _render_alert("โหลด Mapping และปรับใช้กับ Sheet ที่ตรวจพบแล้ว", "success")
            except Exception as exc:  # pragma: no cover - UI guardrail
                st.error(f"ไม่สามารถโหลดไฟล์ Mapping ได้: {exc}")

        if mapping_preset_upload is not None:
            try:
                preset_upload_bytes = mapping_preset_upload.getvalue()
                preset_upload_identity = (
                    mapping_preset_upload.name,
                    hashlib.sha256(preset_upload_bytes).hexdigest(),
                )
                if st.session_state.get("tmc_mapping_preset_upload_identity") != preset_upload_identity:
                    if st.session_state.get("tmc_mapping_preset_ignored_identity") != preset_upload_identity:
                        st.session_state.pop("tmc_mapping_preset_ignored_identity", None)
                        loaded_preset = load_mapping_preset(preset_upload_bytes)
                        apply_result = apply_mapping_preset_to_detected_sheets(loaded_preset, detected_sheet_names)
                        default_mapping = apply_result.mapping
                        detected_scheme = detect_mapping_preset_scheme(loaded_preset)
                        _set_current_mapping_scheme(detected_scheme)
                        st.session_state["tmc_loaded_mapping_code_scheme"] = detected_scheme
                        st.session_state["mapping_table"] = default_mapping.to_dict("records")
                        st.session_state["mapping_editor_version"] = int(st.session_state.get("mapping_editor_version", 0) or 0) + 1
                        st.session_state["tmc_mapping_table_from_session"] = False
                        _set_mapping_source(MAPPING_SOURCE_MAPPING_PRESET)
                        st.session_state["tmc_mapping_preset_upload_identity"] = preset_upload_identity
                        st.session_state["tmc_mapping_preset_apply_info"] = {
                            "matched": apply_result.matched_sheet_count,
                            "missing": apply_result.missing_detected_sheet_count,
                            "extra": apply_result.extra_preset_row_count,
                            "scheme": detected_scheme,
                        }
                        st.session_state["tmc_mapping_preset_warnings"] = list(loaded_preset.warnings)
            except (MappingPresetError, ValueError) as exc:
                st.error(f"ไม่สามารถเปิด Mapping Preset ได้: {exc}")
        preset_info = st.session_state.get("tmc_mapping_preset_apply_info")
        if preset_info:
            _render_alert("โหลด Mapping Preset สำเร็จ", "success")
            _render_alert(
                f"จับคู่ Sheet ได้ {preset_info.get('matched', 0)} รายการ; "
                f"ยังต้องตรวจสอบ Sheet ที่พบใหม่ {preset_info.get('missing', 0)} รายการ; "
                f"แถวจาก Preset ที่ไม่พบใน workbook นี้ {preset_info.get('extra', 0)} รายการ",
                "info",
            )
        for warning_message in st.session_state.get("tmc_mapping_preset_warnings", []):
            _render_alert(warning_message, "warning")

        mapping_scheme = _current_mapping_scheme()
        mapping_editor_version = int(st.session_state.get("mapping_editor_version", 0) or 0)
        mapping_view = str(st.session_state.get("mapping_editor_view_mode") or "Basic")
        editor_key = _mapping_editor_key(mapping_view, mapping_editor_version)
        if mapping_view in {"Basic", "Advanced"}:
            default_mapping = _mapping_from_editor_widget_state(
                default_mapping,
                mapping_view,
                mapping_scheme,
                editor_key,
                apply_basic_widgets=True,
            )
        if _is_v2_scheme(mapping_scheme):
            default_mapping = normalize_approach_movement_mapping(default_mapping)
        if _basic_mapping_widget_state_exists(
            mapping_editor_version,
            st.session_state,
        ):
            st.session_state["mapping_table"] = default_mapping.to_dict("records")
            st.session_state["tmc_mapping_table_from_session"] = False
            if _mapping_source() == MAPPING_SOURCE_DEFAULT_PREVIEW:
                _set_mapping_source(MAPPING_SOURCE_USER_EDITOR)
        scheme_validation_issues = validate_mapping_scheme(default_mapping, mapping_scheme)
        process_block_reason = _single_file_processing_block_reason(mapping_scheme)
        mapping_issues = _single_file_mapping_issues(detected_sheet_names, default_mapping, mapping_scheme)
        mapping_counts = _mapping_workspace_counts(default_mapping, detected_sheet_names)
        for issue in scheme_validation_issues:
            _render_alert(issue, "warning")
        _render_metric_strip(
            [
                ("ไฟล์สำรวจ", "โหลดแล้ว", "", uploaded_file.name, "พร้อม"),
                ("Sheet ที่พบ", mapping_counts["detected_sheets"], "sheet", "ตรวจจาก workbook", "พร้อม"),
                ("แถว Mapping", mapping_counts["rows"], "แถว", "ข้อมูลหลักสำหรับประมวลผล", "พร้อม" if mapping_counts["rows"] else "ต้องตรวจสอบ"),
                ("Movement ที่ใช้", mapping_counts["included"], "แถว", f"ไม่รวม {mapping_counts['excluded']:,} แถว", "พร้อม" if mapping_counts["included"] else "ต้องตรวจสอบ"),
                ("รวมหลาย source", mapping_counts["duplicate_movements"], "movement", "อนุญาตสำหรับ aggregation", "ข้อมูล" if mapping_counts["duplicate_movements"] else "พร้อม"),
                ("สถานะ Mapping", "พร้อม" if mapping_issues.empty else "ต้องตรวจสอบ", "", "ประมวลผลได้" if mapping_issues.empty else f"{len(mapping_issues):,} รายการ", "พร้อม" if mapping_issues.empty else "ต้องตรวจสอบ"),
            ],
            columns=6,
        )

        if process_block_reason:
            _render_alert(process_block_reason, "warning")
        elif mapping_issues.empty and not scheme_validation_issues:
            _render_action_hint("Mapping is ready. Continue to Analyze to run Analyze TMC.")
        else:
            _render_alert("กรุณาตรวจสอบ Mapping ก่อนประมวลผล", "warning")

        for warning_message in mapping_control_warnings(default_mapping, mapping_scheme):
            _render_alert(_thai_mapping_control_warning(warning_message), "warning")
        movement_code_options = _movement_code_options_for_scheme(default_mapping, mapping_scheme)
        with st.expander("Sheet ทิศทางที่ตรวจพบ", expanded=False):
            st.dataframe(preview_summary, width="stretch")
        _render_section_header("ตาราง Mapping", "ตารางนี้เป็นข้อมูลหลักสำหรับการจับคู่ movement")
        mapping_view = st.radio(
            "มุมมองตาราง Mapping",
            options=["Basic", "Advanced"],
            horizontal=True,
            key="mapping_editor_view_mode",
            help="Basic แสดงคอลัมน์ที่ใช้บ่อย ส่วน Advanced แสดงคอลัมน์เสริมและเชิงเทคนิค",
        )
        if mapping_view != "Basic":
            _render_mapping_scheme_status(mapping_scheme)
        editor_key = _mapping_editor_key(mapping_view, mapping_editor_version)
        editor_frame = _mapping_editor_frame(default_mapping, mapping_view, mapping_scheme)
        use_basic_controls = mapping_view == "Basic"
        if use_basic_controls:
            edited_frame = _render_basic_mapping_editor(
                editor_frame,
                movement_code_options,
                mapping_editor_version,
                movement_code_scheme=mapping_scheme,
            )
        else:
            mapping = st.data_editor(
                editor_frame,
                width="stretch",
                num_rows="dynamic",
                column_config=_mapping_editor_column_config(mapping_scheme, movement_code_options),
                key=editor_key,
            )
            edited_frame = pd.DataFrame(mapping)
            edited_frame = _apply_mapping_editor_widget_state(edited_frame, st.session_state.get(editor_key))
        editor_changed = _mapping_editor_changed(editor_frame, edited_frame)
        mapping = _merge_mapping_editor_result(default_mapping, edited_frame, mapping_scheme)
        if _mapping_rows_are_committed() or editor_changed:
            st.session_state["mapping_table"] = mapping.to_dict("records")
            st.session_state["tmc_mapping_table_from_session"] = False
            if editor_changed and _mapping_source() == MAPPING_SOURCE_DEFAULT_PREVIEW:
                _set_mapping_source(MAPPING_SOURCE_USER_EDITOR)
        mapping_counts = _mapping_workspace_counts(mapping, detected_sheet_names)

        aggregation_preview = _mapping_aggregation_preview(mapping)
        if not aggregation_preview.empty:
            _render_alert("พบ movement ที่รวมจากหลาย source stream", "info")
            with st.expander("Audit: movement ที่รวมจากหลาย source stream", expanded=False):
                st.dataframe(aggregation_preview, width="stretch", hide_index=True)
                for aggregation_message in movement_aggregation_messages(mapping):
                    st.caption(_thai_aggregation_message(aggregation_message))

        mapping_scheme = _current_mapping_scheme()
        scheme_validation_issues = validate_mapping_scheme(mapping, mapping_scheme)
        process_block_reason = _single_file_processing_block_reason(mapping_scheme)
        mapping_issues = _single_file_mapping_issues(detected_sheet_names, mapping, mapping_scheme)
        if process_block_reason:
            _render_alert(process_block_reason, "warning")
        elif _is_v2_scheme(mapping_scheme) and not scheme_validation_issues:
            _render_alert("Mapping approach_movement พร้อมใช้งานสำหรับ single-file workflow", "success")
        elif mapping_issues.empty and not scheme_validation_issues:
            _render_alert("Mapping พร้อมใช้งาน", "success")
        else:
            _render_alert("มีรายการที่ต้องตรวจสอบก่อนประมวลผล", "warning")
            with st.expander("รายการที่ต้องตรวจสอบ", expanded=True):
                st.dataframe(_mapping_issue_display(mapping_issues), width="stretch")
        if mapping_counts["blank_source_stream"]:
            _render_alert(f"พบ source_stream ว่าง {mapping_counts['blank_source_stream']:,} แถว ระบบจะใช้ค่าเริ่มต้นตามพฤติกรรมเดิมเมื่อรวมข้อมูล", "info")
        if mapping_counts["excluded"]:
            _render_alert(f"มีแถวที่ไม่แสดงในรายงาน {mapping_counts['excluded']:,} แถว", "info")

        _sync_single_workflow_from_state(
            source_bytes=file_bytes if uploaded_file is not None else None,
            source_file_name=uploaded_file.name if uploaded_file is not None else None,
            export_mode=export_mode,
        )
def render_single_analyze(*, context: WorkflowContext) -> None:
    """Render the Single Analyze stage."""
    AM_WINDOW = context.operations.AM_WINDOW
    DEFAULT_PEAK_MODE = context.operations.DEFAULT_PEAK_MODE
    PEAK_MODE_OPTIONS = context.operations.PEAK_MODE_OPTIONS
    PM_WINDOW = context.operations.PM_WINDOW
    WORKFLOW_SINGLE_MODE = context.operations.WORKFLOW_SINGLE_MODE
    WorkflowReadiness = context.operations.WorkflowReadiness
    _coerce_setup_time = context.operations._coerce_setup_time
    _confirmed_peaks_from_state = context.operations._confirmed_peaks_from_state
    _current_mapping_scheme = context.operations._current_mapping_scheme
    _flash_and_rerun = context.operations._flash_and_rerun
    _pce_override_summary = context.operations._pce_override_summary
    _process_single_file_for_ui = context.operations._process_single_file_for_ui
    _record_workflow_state = context.operations._record_workflow_state
    _render_action_hint = context.operations._render_action_hint
    _render_alert = context.operations._render_alert
    _render_pce_factor_editor = context.operations._render_pce_factor_editor
    _render_readiness_checklist = context.operations._render_readiness_checklist
    _render_section_header = context.operations._render_section_header
    _render_status_chip = context.operations._render_status_chip
    _single_effective_peak_state = context.operations._single_effective_peak_state
    _single_file_mapping_issues = context.operations._single_file_mapping_issues
    _single_file_processing_block_reason = context.operations._single_file_processing_block_reason
    _single_workflow_revisions = context.operations._single_workflow_revisions
    _sync_workflow_after_pce_editor = context.operations._sync_workflow_after_pce_editor
    _time_from_text = context.operations._time_from_text
    _workflow_state_for_mode = context.operations._workflow_state_for_mode
    analysis_block_reason = context.analysis_block_reason
    analysis_ready = context.analysis_ready
    analysis_stale = context.analysis_stale
    analyze_tmc = context.analyze_tmc
    build_setup_for_processing = context.operations.build_setup_for_processing
    detected_sheet_names = context.detected_sheet_names
    exc = context.exc
    export_mode = context.export_mode
    file_bytes = context.file_bytes
    has_overrides = context.has_overrides
    mapping = context.mapping
    mapping_issues = context.mapping_issues
    mapping_scheme = context.mapping_scheme
    name = context.name
    override_text = context.override_text
    parsed = context.parsed
    parsed_details = context.parsed_details
    peak_mode = context.peak_mode
    peak_mode_default = context.peak_mode_default
    peak_windows = context.peak_windows
    period_cols = context.period_cols
    process_block_reason = context.process_block_reason
    processed_revisions = context.processed_revisions
    raw_sheets = context.raw_sheets
    result = context.result
    scheme_validation_issues = context.scheme_validation_issues
    selected_pce_factors = context.selected_pce_factors
    set_active_tab = context.operations.set_active_tab
    setup = context.setup
    uploaded_file = context.uploaded_file
    validate_mapping_scheme = context.operations.validate_mapping_scheme
    workflow_state = context.workflow_state

    _render_section_header("Analyze", "Configure analysis settings, review blockers, and run Analyze TMC.")
    mapping_scheme = _current_mapping_scheme()
    scheme_validation_issues = validate_mapping_scheme(mapping, mapping_scheme)
    mapping_issues = _single_file_mapping_issues(detected_sheet_names, mapping, mapping_scheme)
    process_block_reason = _single_file_processing_block_reason(mapping_scheme)
    workflow_state = _workflow_state_for_mode(WORKFLOW_SINGLE_MODE)
    analysis_ready = bool(workflow_state and workflow_state.readiness.analysis)
    analysis_stale = bool(st.session_state.get("tmc_pce_results_stale")) or bool(
        workflow_state and workflow_state.readiness.mapping and not analysis_ready and st.session_state.get("tmc_processed")
    )
    _render_readiness_checklist(
        [
            ("Source workbook", uploaded_file is not None, uploaded_file.name if uploaded_file is not None else "Upload in Data"),
            ("Mapping", mapping_issues.empty and not scheme_validation_issues and not process_block_reason, "Ready" if mapping_issues.empty and not scheme_validation_issues and not process_block_reason else "Fix Mapping first"),
            ("Analysis result", analysis_ready and not analysis_stale, "Re-analysis required" if analysis_stale else "Not analyzed" if not analysis_ready else "Ready"),
        ]
    )
    if analysis_stale:
        _render_alert("Analysis is stale. Update the settings below and run Analyze TMC again.", "warning")
    elif workflow_state and workflow_state.readiness.analysis:
        _render_alert("Analysis result is current for the active source, Mapping, and settings.", "success")

    with st.container(border=True):
        _render_section_header("Analysis settings", "Peak search windows and PCE factors are part of Analyze.")
        survey_period = st.text_input("Survey period", key="survey_period_input")
        peak_mode_default = str(st.session_state.get("peak_mode_select") or DEFAULT_PEAK_MODE)
        if peak_mode_default not in PEAK_MODE_OPTIONS:
            peak_mode_default = DEFAULT_PEAK_MODE
            st.session_state["peak_mode_select"] = peak_mode_default
        peak_mode = st.selectbox(
            "Peak calculation mode",
            options=PEAK_MODE_OPTIONS,
            index=PEAK_MODE_OPTIONS.index(peak_mode_default),
            key="peak_mode_select",
        )
        period_cols = st.columns(4)
        am_peak_window_start = period_cols[0].time_input(
            "AM window start",
            value=_coerce_setup_time(st.session_state["am_peak_window_start_input"], _time_from_text(AM_WINDOW[0])),
            step=900,
            key="am_peak_window_start_input",
        )
        am_peak_window_end = period_cols[1].time_input(
            "AM window end",
            value=_coerce_setup_time(st.session_state["am_peak_window_end_input"], _time_from_text(AM_WINDOW[1])),
            step=900,
            key="am_peak_window_end_input",
        )
        pm_peak_window_start = period_cols[2].time_input(
            "PM window start",
            value=_coerce_setup_time(st.session_state["pm_peak_window_start_input"], _time_from_text(PM_WINDOW[0])),
            step=900,
            key="pm_peak_window_start_input",
        )
        pm_peak_window_end = period_cols[3].time_input(
            "PM window end",
            value=_coerce_setup_time(st.session_state["pm_peak_window_end_input"], _time_from_text(PM_WINDOW[1])),
            step=900,
            key="pm_peak_window_end_input",
        )
        selected_pce_factors = _render_pce_factor_editor()
        _sync_workflow_after_pce_editor(
            is_single_file_mode=True,
            source_bytes=file_bytes if uploaded_file is not None else None,
            source_file_name=uploaded_file.name if uploaded_file is not None else None,
            export_mode=export_mode,
        )
        has_overrides, override_text = _pce_override_summary(selected_pce_factors)
        _render_status_chip("User-adjusted PCE" if has_overrides else "Default PCE factors", "warning" if has_overrides else "success")
        if has_overrides:
            st.caption(override_text)

    setup = build_setup_for_processing(uploaded_file.name if uploaded_file is not None else "")
    setup["movement_code_scheme"] = _current_mapping_scheme()
    peak_mode = str(setup.get("peak_mode", DEFAULT_PEAK_MODE) or DEFAULT_PEAK_MODE)
    peak_windows = {
        "AM": (setup["am_peak_window_start"], setup["am_peak_window_end"]),
        "PM": (setup["pm_peak_window_start"], setup["pm_peak_window_end"]),
    }

    analysis_block_reason = ""
    if uploaded_file is None:
        analysis_block_reason = "Upload a source workbook in Data first."
    elif process_block_reason:
        analysis_block_reason = process_block_reason
    elif not mapping_issues.empty or scheme_validation_issues:
        analysis_block_reason = "Resolve Mapping issues before analysis."
    _render_action_hint(analysis_block_reason or "Ready to analyze the current source and Mapping.")
    analyze_tmc = st.button(
        "Analyze TMC",
        type="primary",
        disabled=bool(analysis_block_reason),
        key="analyze_tmc_stage",
    )
    if analyze_tmc:
        try:
            raw_sheets = {name: parsed.data for name, parsed in parsed_details.items()}
            with st.spinner("Analyzing TMC..."):
                result = _process_single_file_for_ui(
                    raw_sheets=raw_sheets,
                    mapping=mapping,
                    setup=setup,
                    detected_sheets=detected_sheet_names,
                    peak_mode=peak_mode,
                    peak_windows=peak_windows,
                    pce_factors=selected_pce_factors,
                )
        except Exception as exc:  # pragma: no cover - UI guardrail
            st.error(f"Analysis failed: {exc}")
        else:
            st.session_state["tmc_processed"] = {
                "result": result,
                "mapping": mapping,
                "setup": setup,
                "peak_windows": peak_windows,
                "pce_factors": selected_pce_factors,
            }
            st.session_state.pop("tmc_output", None)
            st.session_state.pop("tmc_pce_results_stale", None)
            processed_revisions = _single_workflow_revisions(
                source_bytes=file_bytes if uploaded_file is not None else None,
                source_file_name=uploaded_file.name if uploaded_file is not None else None,
                mapping=mapping,
                pce_factors=selected_pce_factors,
                peak_mode=peak_mode,
                peak_windows=peak_windows,
                movement_code_scheme=_current_mapping_scheme(),
                setup=setup,
                export_mode=export_mode,
                confirmed_peaks=_confirmed_peaks_from_state(),
                analysis_present=True,
            )
            _record_workflow_state(
                WORKFLOW_SINGLE_MODE,
                processed_revisions,
                WorkflowReadiness(
                    source=uploaded_file is not None,
                    mapping=bool(mapping.to_dict("records")),
                    analysis=True,
                    review=bool(_single_effective_peak_state().get("ready")),
                    export=False,
                ),
            )
            set_active_tab("Review")
            _flash_and_rerun("Analyze TMC completed. Review the AM/PM Peak result next.")
def render_single_review(*, context: WorkflowContext) -> None:
    """Render the Single Review stage."""
    DEFAULT_TEMPLATE_MAP_PATH = context.operations.DEFAULT_TEMPLATE_MAP_PATH
    DEFAULT_TEMPLATE_PATH = context.operations.DEFAULT_TEMPLATE_PATH
    Path = context.operations.Path
    TEMPLATE_VERSION = context.operations.TEMPLATE_VERSION
    _confirm_single_peak_review = context.operations._confirm_single_peak_review
    _confirmed_peaks_from_state = context.operations._confirmed_peaks_from_state
    _flash_and_rerun = context.operations._flash_and_rerun
    _format_display_columns = context.operations._format_display_columns
    _hourly_interval_options = context.operations._hourly_interval_options
    _interval_total_pcu = context.operations._interval_total_pcu
    _is_v2_result = context.operations._is_v2_result
    _peak_period_text = context.operations._peak_period_text
    _peak_value_payload = context.operations._peak_value_payload
    _peak_values_complete = context.operations._peak_values_complete
    _qc_display_frame = context.operations._qc_display_frame
    _qc_severity_counts = context.operations._qc_severity_counts
    _render_action_hint = context.operations._render_action_hint
    _render_alert = context.operations._render_alert
    _render_empty_state = context.operations._render_empty_state
    _render_excel_com_status = context.operations._render_excel_com_status
    _render_hourly_pcu_line_chart = context.operations._render_hourly_pcu_line_chart
    _render_metric_strip = context.operations._render_metric_strip
    _render_peak_card = context.operations._render_peak_card
    _render_qc_status = context.operations._render_qc_status
    _render_section_header = context.operations._render_section_header
    _render_template_audit_notes = context.operations._render_template_audit_notes
    _selected_interval = context.operations._selected_interval
    _sync_single_workflow_from_state = context.operations._sync_single_workflow_from_state
    _v2_movement_code_reference_frame = context.operations._v2_movement_code_reference_frame
    am_default = context.am_default
    am_end = context.am_end
    am_index = context.am_index
    am_pcu = context.am_pcu
    am_peak_label = context.am_peak_label
    am_start = context.am_start
    audit_frame = context.audit_frame
    build_v2_movement_diagram_data = context.operations.build_v2_movement_diagram_data
    confirm_cols = context.confirm_cols
    confirm_review = context.confirm_review
    confirmed_am_end = context.confirmed_am_end
    confirmed_am_label = context.confirmed_am_label
    confirmed_am_start = context.confirmed_am_start
    confirmed_pm_end = context.confirmed_pm_end
    confirmed_pm_label = context.confirmed_pm_label
    confirmed_pm_start = context.confirmed_pm_start
    debug = context.debug
    diagram_data = context.diagram_data
    draft_confirmed = context.draft_confirmed
    excel_com_status = context.excel_com_status
    export_mode = context.export_mode
    file_bytes = context.file_bytes
    format_count = context.operations.format_count
    format_pcu = context.operations.format_pcu
    hourly_movement = context.hourly_movement
    interval_options = context.interval_options
    label = context.label
    loaded_am_label = context.loaded_am_label
    loaded_confirmed_peaks = context.loaded_confirmed_peaks
    loaded_pm_label = context.loaded_pm_label
    mapping_df = context.mapping_df
    movement_aggregation_audit = context.operations.movement_aggregation_audit
    option_labels = context.option_labels
    output = context.output
    parsed = context.parsed
    parsed_details = context.parsed_details
    pce_results_stale = context.pce_results_stale
    peak_display = context.peak_display
    pm_default = context.pm_default
    pm_end = context.pm_end
    pm_index = context.pm_index
    pm_pcu = context.pm_pcu
    pm_peak_label = context.pm_peak_label
    pm_start = context.pm_start
    preview = context.preview
    preview_summary = context.preview_summary
    previews = context.previews
    previous_confirmed_state = context.previous_confirmed_state
    qc_counts = context.qc_counts
    qc_display = context.qc_display
    result = context.result
    sheet_name = context.sheet_name
    total_qc = context.total_qc
    uploaded_file = context.uploaded_file
    value = context.value

    _render_section_header("ตรวจ Peak", "ตรวจสอบรูปแบบปริมาณจราจรรายชั่วโมง และกำหนดช่วง AM/PM Peak สำหรับใช้ในรายงาน")
    if pce_results_stale:
        _render_alert("ค่า PCE เปลี่ยนหลังจากประมวลผลแล้ว กรุณาประมวลผลใหม่ก่อนตรวจ Peak หรือส่งออกรายงาน", "warning")
    if result is None:
        _render_empty_state(
            "กรุณาประมวลผลข้อมูลก่อนตรวจ Peak",
            "เมื่อประมวลผลแล้ว ระบบจะแสดงกราฟ PCU รายชั่วโมงและตัวเลือกยืนยัน AM/PM Peak",
        )
    else:
        am_start, am_end, am_pcu = _peak_period_text(result.peaks, "AM")
        pm_start, pm_end, pm_pcu = _peak_period_text(result.peaks, "PM")
        _render_section_header("กราฟ PCU รายชั่วโมง", "ตรวจแนวโน้มปริมาณรวมก่อนเลือกช่วง Peak ที่ใช้เป็นค่าหลัก")
        _render_hourly_pcu_line_chart(
            hourly_movement,
            f"{am_start}-{am_end}" if am_start and am_end else "",
            f"{pm_start}-{pm_end}" if pm_start and pm_end else "",
        )

        interval_options = _hourly_interval_options(hourly_movement, result.peaks)
        if interval_options:
            option_labels = [label for label, _, _ in interval_options]
            am_default = f"{am_start}-{am_end}" if am_start and am_end else option_labels[0]
            pm_default = f"{pm_start}-{pm_end}" if pm_start and pm_end else option_labels[min(1, len(option_labels) - 1)]
            loaded_confirmed_peaks = st.session_state.get("tmc_loaded_confirmed_peaks") or {}
            loaded_am_label = f"{loaded_confirmed_peaks.get('am_peak_start', '')}-{loaded_confirmed_peaks.get('am_peak_end', '')}".strip("-")
            loaded_pm_label = f"{loaded_confirmed_peaks.get('pm_peak_start', '')}-{loaded_confirmed_peaks.get('pm_peak_end', '')}".strip("-")
            if loaded_am_label in option_labels and not st.session_state.get("am_peak_period_select"):
                am_default = loaded_am_label
            if loaded_pm_label in option_labels and not st.session_state.get("pm_peak_period_select"):
                pm_default = loaded_pm_label
            am_index = option_labels.index(am_default) if am_default in option_labels else 0
            pm_index = option_labels.index(pm_default) if pm_default in option_labels else min(1, len(option_labels) - 1)

            _render_section_header(
                "กำหนดช่วง Peak",
                "ระบบตรวจจับอัตโนมัติเป็นค่าแนะนำ ผู้ตรวจกำหนดช่วงที่จะใช้ในรายงาน",
            )
            confirm_cols = st.columns(2)
            with confirm_cols[0]:
                _render_peak_card("AM Peak · ระบบตรวจจับอัตโนมัติ", f"{am_start}-{am_end}" if am_start and am_end else "", am_pcu, "auto_suggested")
                am_peak_label = st.selectbox("ช่วงที่กำหนด AM", option_labels, index=am_index, key="am_peak_period_select")
            with confirm_cols[1]:
                _render_peak_card("PM Peak · ระบบตรวจจับอัตโนมัติ", f"{pm_start}-{pm_end}" if pm_start and pm_end else "", pm_pcu, "auto_suggested")
                pm_peak_label = st.selectbox("ช่วงที่กำหนด PM", option_labels, index=pm_index, key="pm_peak_period_select")
            confirmed_am_start, confirmed_am_end = _selected_interval(interval_options, am_peak_label)
            confirmed_pm_start, confirmed_pm_end = _selected_interval(interval_options, pm_peak_label)
            draft_confirmed = (confirmed_am_start, confirmed_am_end, confirmed_pm_start, confirmed_pm_end)
            previous_confirmed_state = _peak_value_payload(_confirmed_peaks_from_state())
            if _peak_values_complete(previous_confirmed_state):
                _render_alert(
                    "Draft Peak changes are not applied until you press Confirm Peak Review.",
                    "warning"
                    if draft_confirmed
                    != (
                        previous_confirmed_state.get("am_peak_start"),
                        previous_confirmed_state.get("am_peak_end"),
                        previous_confirmed_state.get("pm_peak_start"),
                        previous_confirmed_state.get("pm_peak_end"),
                    )
                    else "info",
                )
            else:
                _render_alert("Suggested Peaks are drafts until you press Confirm Peak Review.", "info")
            confirm_review = st.button(
                "Confirm Peak Review",
                type="primary",
                disabled=not all(draft_confirmed),
                key="confirm_single_peak_review",
            )
            if confirm_review:
                _confirm_single_peak_review(
                    (confirmed_am_start, confirmed_am_end),
                    (confirmed_pm_start, confirmed_pm_end),
                )
                _sync_single_workflow_from_state(
                    source_bytes=file_bytes if uploaded_file is not None else None,
                    source_file_name=uploaded_file.name if uploaded_file is not None else None,
                    export_mode=export_mode,
                )
                _flash_and_rerun("Peak Review confirmed. Analysis remains current.")
            confirmed_am_start = previous_confirmed_state.get("am_peak_start", "")
            confirmed_am_end = previous_confirmed_state.get("am_peak_end", "")
            confirmed_pm_start = previous_confirmed_state.get("pm_peak_start", "")
            confirmed_pm_end = previous_confirmed_state.get("pm_peak_end", "")
            confirmed_am_label = f"{confirmed_am_start}-{confirmed_am_end}" if confirmed_am_start and confirmed_am_end else ""
            confirmed_pm_label = f"{confirmed_pm_start}-{confirmed_pm_end}" if confirmed_pm_start and confirmed_pm_end else ""
            _render_metric_strip(
                [
                    ("AM Peak", confirmed_am_label or "-", "", "ช่วงที่กำหนด", "กำหนดแล้ว" if confirmed_am_label else "รอตรวจสอบ"),
                    ("PM Peak", confirmed_pm_label or "-", "", "ช่วงที่กำหนด", "กำหนดแล้ว" if confirmed_pm_label else "รอตรวจสอบ"),
                    ("AM PCU", _interval_total_pcu(hourly_movement, confirmed_am_label) or am_pcu or "-", "PCU", "Peak PCU"),
                    ("PM PCU", _interval_total_pcu(hourly_movement, confirmed_pm_label) or pm_pcu or "-", "PCU", "Peak PCU"),
                ],
                columns=4,
            )
            _render_action_hint("ใช้ช่วงนี้เป็นค่าหลักสำหรับรายงาน")
            if all([confirmed_am_start, confirmed_am_end, confirmed_pm_start, confirmed_pm_end]):
                _render_alert("กำหนดช่วง Peak แล้ว พร้อมส่งออก", "success")
            else:
                _render_alert("กรุณากำหนด AM Peak และ PM Peak ก่อนส่งออก", "warning")
        else:
            _render_alert("ไม่มีช่วงเวลารายชั่วโมงสำหรับกำหนด Peak", "warning")

        _render_section_header("สรุปทางเทคนิค", "แสดงเฉพาะค่าที่มีจากผลประมวลผลปัจจุบัน")
        _render_metric_strip(
            [
                ("จำนวนแถว", format_count(len(result.normalized)), "แถว", "normalized"),
                ("จำนวนรถรวม", format_count(result.normalized["count"].sum()) if not result.normalized.empty else "0", "คัน", ""),
                ("PCU รวม", format_pcu(result.normalized["pcu"].sum()) if not result.normalized.empty else "0", "PCU", ""),
                ("QC", format_count(len(result.qc)), "", "ประเด็น"),
            ],
            columns=4,
        )
        _render_qc_status(result.qc)

        with st.expander("ตารางปริมาณจราจรแยกตามทิศทาง", expanded=False):
            st.dataframe(_format_display_columns(hourly_movement), width="stretch", hide_index=True)

    _render_section_header(
        "ตรวจสอบข้อมูล",
        "ตรวจสอบ QC, ข้อมูลที่ประมวลผลแล้ว และรายการ Audit ก่อนนำผลไปใช้ในรายงาน",
    )
    if result is None:
        _render_empty_state(
            "กรุณาประมวลผลข้อมูลก่อนตรวจสอบรายละเอียด",
            "QC, ตารางข้อมูล และรายการ Audit จะแสดงหลังมีผลประมวลผลแล้ว",
        )
    else:
        qc_counts = _qc_severity_counts(result.qc)
        total_qc = sum(qc_counts.values())
        _render_metric_strip(
            [
                ("QC ผิดพลาด", f"{qc_counts['error']:,}", "รายการ", "", "error" if qc_counts["error"] else "พร้อม"),
                ("QC เตือน", f"{qc_counts['warning']:,}", "รายการ", "", "warning" if qc_counts["warning"] else "พร้อม"),
                ("QC ข้อมูล", f"{qc_counts['info']:,}", "รายการ", "", "info" if qc_counts["info"] else "พร้อม"),
                ("QC รวม", f"{total_qc:,}", "รายการ", "", "พร้อม" if not total_qc else "ต้องตรวจ"),
            ],
            columns=4,
        )
        if qc_counts["error"]:
            _render_alert("พบ QC error กรุณาตรวจรายละเอียดก่อนส่งออก", "error")
        elif qc_counts["warning"]:
            _render_alert("พบ QC warning ยังส่งออกได้ แต่ควรตรวจหมายเหตุก่อน", "warning")
        elif qc_counts["info"]:
            _render_alert("มีหมายเหตุ QC info สำหรับตรวจสอบ", "info")
        else:
            _render_alert("ไม่พบประเด็น QC", "success")

        _render_section_header("รายละเอียด QC", "รายการตรวจสอบที่ใช้ประกอบการพิจารณาก่อนส่งออกรายงาน")
        qc_display = _qc_display_frame(result.qc, thai_labels=True)
        if qc_display.empty:
            _render_alert("ไม่พบรายการ QC ที่ต้องตรวจสอบ", "success")
        elif len(qc_display) > 25:
            st.dataframe(qc_display.head(25), width="stretch", hide_index=True)
            with st.expander("รายการ QC ทั้งหมด", expanded=False):
                st.dataframe(qc_display, width="stretch", hide_index=True)
        else:
            st.dataframe(qc_display, width="stretch", hide_index=True)

        with st.expander("Normalized Data", expanded=False):
            st.dataframe(_format_display_columns(result.normalized.head(1000)), width="stretch", hide_index=True)
        with st.expander("Hourly Movement PCU", expanded=False):
            if not hourly_movement.empty:
                st.dataframe(_format_display_columns(hourly_movement), width="stretch", hide_index=True)
            if not result.hourly.empty:
                st.caption("Hourly raw summary")
                st.dataframe(_format_display_columns(result.hourly), width="stretch", hide_index=True)
        with st.expander("Movement Summary", expanded=False):
            st.dataframe(_format_display_columns(result.movement), width="stretch", hide_index=True)
        if _is_v2_result(result):
            with st.expander("Movement Diagram Data", expanded=False):
                diagram_data = build_v2_movement_diagram_data(
                    movement_summary=result.movement,
                    hourly_movement_pcu=hourly_movement,
                    peaks=result.peaks,
                )
                st.dataframe(_format_display_columns(diagram_data), width="stretch", hide_index=True)
            with st.expander("Movement Code Reference", expanded=False):
                st.dataframe(_v2_movement_code_reference_frame(), width="stretch", hide_index=True)
        with st.expander("Peak / PHF Data", expanded=False):
            peak_display = _format_display_columns(result.peaks)
            if "peak_start" in peak_display.columns:
                peak_display["peak_start"] = peak_display["peak_start"].map(lambda value: str(value)[:5] if str(value) else "-")
            if "peak_end" in peak_display.columns:
                peak_display["peak_end"] = peak_display["peak_end"].map(lambda value: str(value)[:5] if str(value) else "-")
            st.dataframe(peak_display, width="stretch", hide_index=True)
        with st.expander("Movement Aggregation Audit", expanded=False):
            st.caption("ตารางตรวจสอบ source movement, source stream และ output movement ที่ใช้รวมค่าในรายงาน")
            audit_frame = movement_aggregation_audit(result.normalized, mapping_df)
            st.dataframe(_format_display_columns(audit_frame), width="stretch", hide_index=True)
        with st.expander("รายละเอียดการอ่านไฟล์และ Parser", expanded=False):
            if uploaded_file is not None:
                st.dataframe(preview_summary, width="stretch")
                for sheet_name, preview in previews.items():
                    st.markdown(f"**{sheet_name}**")
                    st.dataframe(preview, width="stretch")
                for sheet_name, parsed in parsed_details.items():
                    debug = parsed.debug
                    st.markdown(f"**Parser: {sheet_name}**")
                    st.write(
                        {
                            "detected_first_data_row": debug.first_data_row,
                            "detected_time_columns": {
                                "time_start_col": debug.time_start_col,
                                "time_end_col": debug.time_end_col,
                            },
                            "detected_vehicle_class_columns": debug.vehicle_class_columns,
                        }
                    )
                    st.dataframe(parsed.data.head(10), width="stretch")
            else:
                _render_action_hint("ไม่มี Workbook ที่อัปโหลดในรอบการทำงานนี้")
        with st.expander("Export Metadata / Template Diagnostics", expanded=False):
            output = st.session_state.get("tmc_output")
            if output:
                st.write(
                    {
                        "workbook_filename": output.get("workbook_filename", ""),
                        "export_mode": output.get("export_mode", ""),
                        "generated_at": output.get("generated_at"),
                        "template_version": TEMPLATE_VERSION,
                        "template_name": Path(DEFAULT_TEMPLATE_PATH).name,
                        "template_map_name": Path(DEFAULT_TEMPLATE_MAP_PATH).name,
                    }
                )
            else:
                _render_action_hint("ยังไม่มี metadata การส่งออกในรอบนี้")
            _render_template_audit_notes()
            _render_excel_com_status(excel_com_status)
def render_single_export(*, context: WorkflowContext) -> None:
    """Render the Single Export stage."""
    DEFAULT_TEMPLATE_MAP_PATH = context.operations.DEFAULT_TEMPLATE_MAP_PATH
    DEFAULT_TEMPLATE_PATH = context.operations.DEFAULT_TEMPLATE_PATH
    DiagramConfig = context.operations.DiagramConfig
    EXCEL_MIME = context.operations.EXCEL_MIME
    EXCEL_TEMPLATE_EXPORT_MODE = context.operations.EXCEL_TEMPLATE_EXPORT_MODE
    EXPORT_PREFERENCE_ADVANCED = context.operations.EXPORT_PREFERENCE_ADVANCED
    EXPORT_PREFERENCE_STANDARD = context.operations.EXPORT_PREFERENCE_STANDARD
    MOVEMENT_SCHEME_V2 = context.operations.MOVEMENT_SCHEME_V2
    PACKAGE_MIME = context.operations.PACKAGE_MIME
    PEAK_SELECTION_AUTO = context.operations.PEAK_SELECTION_AUTO
    PNG_MIME = context.operations.PNG_MIME
    PROJECT_SESSION_MIME = context.operations.PROJECT_SESSION_MIME
    Path = context.operations.Path
    STANDARD_REPORT_EXPORT_MODE = context.operations.STANDARD_REPORT_EXPORT_MODE
    TEMPLATE_VERSION = context.operations.TEMPLATE_VERSION
    V2_EXCEL_TEMPLATE_MODE_BLOCK_MESSAGE = context.operations.V2_EXCEL_TEMPLATE_MODE_BLOCK_MESSAGE
    _apply_standard_single_export_mode = context.operations._apply_standard_single_export_mode
    _build_session_from_state = context.operations._build_session_from_state
    _coerce_export_mode = context.operations._coerce_export_mode
    _current_mapping_scheme = context.operations._current_mapping_scheme
    _default_single_export_mode = context.operations._default_single_export_mode
    _export_single_file_for_ui = context.operations._export_single_file_for_ui
    _flash_and_rerun = context.operations._flash_and_rerun
    _hourly_movement_for_ui = context.operations._hourly_movement_for_ui
    _is_v2_result = context.operations._is_v2_result
    _render_action_hint = context.operations._render_action_hint
    _render_alert = context.operations._render_alert
    _render_download_button = context.operations._render_download_button
    _render_effective_peak_export_confirmation = context.operations._render_effective_peak_export_confirmation
    _render_empty_state = context.operations._render_empty_state
    _render_excel_com_status = context.operations._render_excel_com_status
    _render_qc_status = context.operations._render_qc_status
    _render_readiness_checklist = context.operations._render_readiness_checklist
    _render_section_header = context.operations._render_section_header
    _single_effective_peak_state = context.operations._single_effective_peak_state
    _single_export_mode_options = context.operations._single_export_mode_options
    _standard_report_decision = context.operations._standard_report_decision
    _use_excel_native_charts_for_export = context.operations._use_excel_native_charts_for_export
    _use_template_layout_for_export = context.operations._use_template_layout_for_export
    application_analyze_single = context.operations.application_analyze_single
    application_analyze_single_dry_run = context.operations.application_analyze_single_dry_run
    apply_single_export_mode_change = context.operations.apply_single_export_mode_change
    build_export_summary_text = context.operations.build_export_summary_text
    build_mapping_preset = context.operations.build_mapping_preset
    caption_text = context.caption_text
    chart_cols = context.chart_cols
    chart_pngs = context.chart_pngs
    confirmed_hourly_movement = context.confirmed_hourly_movement
    confirmed_periods = context.confirmed_periods
    confirmed_ready = context.confirmed_ready
    confirmed_result = context.confirmed_result
    confirmed_setup = context.confirmed_setup
    create_export_package_zip = context.operations.create_export_package_zip
    create_v2_generated_export_package_zip = context.operations.create_v2_generated_export_package_zip
    detected_sheet_names = context.detected_sheet_names
    diagram_png = context.diagram_png
    download_buffer = context.operations.download_buffer
    east_label = context.east_label
    east_road = context.east_road
    effective_peaks = context.effective_peaks
    exc = context.exc
    excel_com_enabled = context.excel_com_enabled
    excel_com_requested = context.excel_com_requested
    excel_com_status = context.excel_com_status
    export_excel_com_status = context.export_excel_com_status
    export_generated_at = context.export_generated_at
    export_mode = context.export_mode
    export_peak_state = context.export_peak_state
    export_preference = context.export_preference
    export_run = context.export_run
    export_status_col = context.export_status_col
    export_warnings = context.export_warnings
    file_bytes = context.file_bytes
    filename_seed = context.filename_seed
    generate_four_leg_tmc_diagram = context.operations.generate_four_leg_tmc_diagram
    generated_timestamp_text = context.operations.generated_timestamp_text
    is_v2_single_result = context.is_v2_single_result
    mapping_df = context.mapping_df
    mapping_preset_bytes = context.mapping_preset_bytes
    message = context.message
    name = context.name
    north_label = context.north_label
    north_road = context.north_road
    output = context.output
    output_result = context.output_result
    output_setup = context.output_setup
    package_bytes = context.package_bytes
    parsed = context.parsed
    parsed_details = context.parsed_details
    pce_results_stale = context.pce_results_stale
    peak_mode = context.peak_mode
    peak_windows = context.peak_windows
    preview_files = context.preview_files
    previous_export_mode = context.previous_export_mode
    probe_excel_com = context.operations.probe_excel_com
    raw_sheets = context.raw_sheets
    readiness_col = context.readiness_col
    report_chart_pngs = context.operations.report_chart_pngs
    result = context.result
    safe_package_filename = context.operations.safe_package_filename
    safe_project_session_filename = context.operations.safe_project_session_filename
    safe_workbook_filename = context.operations.safe_workbook_filename
    selected_export_mode = context.selected_export_mode
    selected_pce_factors = context.selected_pce_factors
    serialize_mapping_preset = context.operations.serialize_mapping_preset
    session = context.session
    session_bytes = context.session_bytes
    session_filename = context.session_filename
    session_to_json_bytes = context.operations.session_to_json_bytes
    set_active_tab = context.operations.set_active_tab
    setup = context.setup
    show_u_turn = context.show_u_turn
    single_export_options = context.single_export_options
    south_label = context.south_label
    south_road = context.south_road
    standard_decision = context.standard_decision
    standard_mode_changed = context.standard_mode_changed
    summary_text = context.summary_text
    survey_date_text = context.survey_date_text
    survey_period = context.survey_period
    template_compatible = context.template_compatible
    tmc_id = context.tmc_id
    tmc_title = context.tmc_title
    uploaded_file = context.uploaded_file
    use_excel_com_native_charts = context.use_excel_com_native_charts
    use_template_report_layout = context.use_template_report_layout
    vehicle_composition_report = context.operations.vehicle_composition_report
    version_text = context.version_text
    warning = context.warning
    warnings = context.operations.warnings
    west_label = context.west_label
    west_road = context.west_road
    workbook_bytes = context.workbook_bytes

    export_peak_state = _single_effective_peak_state(result if result is not None and not pce_results_stale else None)
    effective_peaks = dict(export_peak_state.get("values") or {})
    confirmed_ready = bool(export_peak_state.get("ready"))
    _render_section_header("ส่งออกรายงาน", "สร้างรายงาน Excel และชุดไฟล์ประกอบสำหรับตรวจสอบย้อนหลัง")
    single_export_options = _single_export_mode_options(excel_com_status)
    previous_export_mode = st.session_state.get("report_export_mode", export_mode)
    standard_decision = None
    export_preference = st.radio(
        "Report outcome",
        options=[EXPORT_PREFERENCE_STANDARD, EXPORT_PREFERENCE_ADVANCED],
        index=0 if st.session_state.get("report_export_preference", EXPORT_PREFERENCE_STANDARD) == EXPORT_PREFERENCE_STANDARD else 1,
        key="report_export_preference_control",
        horizontal=True,
        help="Standard report selects the validated native Excel Template path when available and uses Safe PNG otherwise.",
    )
    st.session_state["report_export_preference"] = export_preference
    if export_preference == EXPORT_PREFERENCE_STANDARD:
        template_compatible = Path(DEFAULT_TEMPLATE_PATH).exists() and Path(DEFAULT_TEMPLATE_MAP_PATH).exists()
        standard_decision = _standard_report_decision(excel_com_status, template_compatible=template_compatible)
        export_mode, standard_mode_changed = _apply_standard_single_export_mode(
            standard_decision,
            previous_export_mode,
        )
        if standard_mode_changed:
            st.rerun()
        use_template_report_layout = bool(standard_decision.use_template_report_layout)
        use_excel_com_native_charts = bool(standard_decision.use_excel_com_native_charts)
        st.info(f"{STANDARD_REPORT_EXPORT_MODE}: {export_mode} selected automatically.")
        if standard_decision.fallback_notice:
            st.warning(standard_decision.fallback_notice)
    else:
        with st.expander("Advanced export options", expanded=True):
            selected_export_mode = st.radio(
                "Backend",
                options=single_export_options,
                index=single_export_options.index(_coerce_export_mode(previous_export_mode, single_export_options, _default_single_export_mode(excel_com_status))),
                key="report_export_mode_control",
                horizontal=True,
                help="Explicit backend selection is intended for advanced users and diagnostics.",
            )
        if apply_single_export_mode_change(selected_export_mode, previous_export_mode):
            st.rerun()
        export_mode = selected_export_mode
        use_template_report_layout = _use_template_layout_for_export(export_mode)
        use_excel_com_native_charts = _use_excel_native_charts_for_export(export_mode, excel_com_status)
    is_v2_single_result = _is_v2_result(result)
    if pce_results_stale:
        _render_alert("ผลลัพธ์เดิมไม่ตรงกับค่า PCE ปัจจุบัน ระบบปิดการส่งออกไว้จนกว่าจะประมวลผลใหม่", "warning")
    _render_action_hint("สร้างรายงานหลังจากประมวลผลและมีช่วงเร่งด่วน AM/PM พร้อมใช้งานแล้ว")
    _render_effective_peak_export_confirmation(result if result is not None else None, export_peak_state)
    export_run = st.button("สร้างรายงาน Excel", type="primary", disabled=not (result is not None and confirmed_ready and not pce_results_stale))

    export_status_col, readiness_col = st.columns([0.9, 1.1])
    with export_status_col:
        with st.container(border=True):
            _render_section_header("โหมดส่งออก", "เลือกวิธีสร้างรายงานจากหน้านี้")
            if export_mode == EXCEL_TEMPLATE_EXPORT_MODE:
                st.markdown('<div class="tmc-mode-note tmc-mode-note-success"><strong>Excel Template Mode</strong> · แนะนำสำหรับรายงานฉบับใช้งานจริง</div>', unsafe_allow_html=True)
                if is_v2_single_result:
                    st.caption(
                        "Excel Template Mode สำหรับ approach_movement ใช้ Excel COM เพื่อรักษากราฟและ diagram จาก template"
                        if use_excel_com_native_charts
                        else V2_EXCEL_TEMPLATE_MODE_BLOCK_MESSAGE
                    )
                else:
                    st.caption("เหมาะสำหรับรายงานฉบับใช้งานจริง รักษา Native Chart และรูปแบบ Excel Template เมื่อ Excel COM พร้อมใช้งาน")
            else:
                st.markdown('<div class="tmc-mode-note tmc-mode-note-warning"><strong>Safe PNG Export Mode</strong> · โหมดสำรอง</div>', unsafe_allow_html=True)
                st.caption("โหมดสำรอง เหมาะสำหรับตรวจร่างหรือกรณี Excel COM ใช้งานไม่ได้")

            if excel_com_status.available:
                version_text = f"Excel version: {excel_com_status.version}" if excel_com_status.version else "พร้อมใช้งาน"
                _render_alert(f"Excel COM พร้อมใช้งาน: {version_text}", "success")
            else:
                _render_alert(f"Excel COM ไม่พร้อมใช้งาน ระบบจะใช้โหมดสำรองแบบ PNG: {excel_com_status.reason}", "warning")
            st.caption("ใช้ปุ่มทดสอบ Excel COM ใน sidebar หากต้องการตรวจสถานะใหม่")
            with st.expander("รายละเอียด Excel COM", expanded=False):
                _render_excel_com_status(excel_com_status)

    with readiness_col:
        with st.container(border=True):
            _render_section_header("ความพร้อมก่อนส่งออก", "รายการตรวจสอบแบบย่อก่อนสร้างรายงาน")
            if result is not None:
                _render_qc_status(result.qc)
            _render_readiness_checklist(
                [
                    ("โหลดไฟล์สำรวจแล้ว", uploaded_file is not None, ""),
                    ("Mapping พร้อมใช้งาน", bool(st.session_state.get("mapping_table")), ""),
                    ("ประมวลผลแล้ว", result is not None, "ค่า PCE เปลี่ยน กรุณาประมวลผลใหม่" if pce_results_stale else ""),
                    ("กำหนดช่วงเร่งด่วน AM/PM แล้ว", confirmed_ready, str(export_peak_state.get("summary_text") or "")),
                    (
                        "Excel COM พร้อมใช้งาน",
                        bool(excel_com_status.available) if export_mode == EXCEL_TEMPLATE_EXPORT_MODE else True,
                        V2_EXCEL_TEMPLATE_MODE_BLOCK_MESSAGE
                        if is_v2_single_result and export_mode == EXCEL_TEMPLATE_EXPORT_MODE and not excel_com_status.available
                        else ("จำเป็นสำหรับ Excel Template Mode" if export_mode == EXCEL_TEMPLATE_EXPORT_MODE else "ไม่จำเป็นในโหมดสำรอง"),
                    ),
                ]
            )

    if export_run:
        set_active_tab("Export")
        excel_com_requested = bool(use_excel_com_native_charts)
        export_excel_com_status = probe_excel_com() if excel_com_requested else None
        excel_com_enabled = bool(export_excel_com_status and export_excel_com_status.available)
        if excel_com_requested and export_excel_com_status is not None and not export_excel_com_status.available:
            st.warning(
                "Excel COM ไม่พร้อมใช้งาน ระบบจะใช้โหมดสำรองแบบ PNG "
                f"สาเหตุ: {export_excel_com_status.reason}. {export_excel_com_status.detail}"
            )

        confirmed_setup = {
            **setup,
            "am_peak_start": effective_peaks.get("am_peak_start", ""),
            "am_peak_end": effective_peaks.get("am_peak_end", ""),
            "pm_peak_start": effective_peaks.get("pm_peak_start", ""),
            "pm_peak_end": effective_peaks.get("pm_peak_end", ""),
            "peak_selection_source": str(export_peak_state.get("source") or PEAK_SELECTION_AUTO),
        }
        confirmed_periods = {
            "AM": (str(effective_peaks.get("am_peak_start", "")), str(effective_peaks.get("am_peak_end", ""))),
            "PM": (str(effective_peaks.get("pm_peak_start", "")), str(effective_peaks.get("pm_peak_end", ""))),
        }
        export_generated_at = generated_timestamp_text()
        try:
            raw_sheets = {name: parsed.data for name, parsed in parsed_details.items()}
            if is_v2_single_result:
                confirmed_result = application_analyze_single_dry_run(
                    raw_sheets=raw_sheets,
                    mapping=mapping_df,
                    setup={**confirmed_setup, "movement_code_scheme": MOVEMENT_SCHEME_V2},
                    detected_sheets=detected_sheet_names,
                    peak_mode=peak_mode,
                    peak_windows=peak_windows,
                    confirmed_peak_periods=confirmed_periods,
                    pce_factors=selected_pce_factors,
                )
                workbook_bytes = _export_single_file_for_ui(
                    result=confirmed_result,
                    mapping=mapping_df,
                    setup=confirmed_setup,
                    export_mode=export_mode,
                    use_template_report_layout=use_template_report_layout,
                    use_excel_com_native_charts=excel_com_enabled,
                    source_file_name=uploaded_file.name if uploaded_file is not None else st.session_state.get("tmc_loaded_source_file_name", ""),
                    generated_at=export_generated_at,
                )
                confirmed_result.workbook_bytes = workbook_bytes
            else:
                with warnings.catch_warnings(record=True) as export_warnings:
                    warnings.simplefilter("always", RuntimeWarning)
                    confirmed_result = application_analyze_single(
                        raw_sheets=raw_sheets,
                        mapping=mapping_df,
                        setup=confirmed_setup,
                        detected_sheets=detected_sheet_names,
                        peak_mode=peak_mode,
                        peak_windows=peak_windows,
                        confirmed_peak_periods=confirmed_periods,
                        pce_factors=selected_pce_factors,
                        generate_workbook=True,
                        use_template_report_layout=use_template_report_layout,
                        use_excel_com_native_charts=excel_com_enabled,
                        export_mode=export_mode,
                        source_file_name=uploaded_file.name if uploaded_file is not None else st.session_state.get("tmc_loaded_source_file_name", ""),
                        generated_at=export_generated_at,
                        export_mode_requested=(STANDARD_REPORT_EXPORT_MODE if standard_decision else export_mode),
                        export_mode_used=export_mode,
                        export_fallback_notice=getattr(standard_decision, "fallback_notice", "") if standard_decision else "",
                    )
                for warning in export_warnings:
                    message = str(warning.message)
                    if excel_com_requested and "Excel COM native-chart export failed after COM was available" in message:
                        st.warning(f"Excel COM ส่งออก Native Chart ไม่สำเร็จ ระบบใช้โหมดสำรองแบบ PNG ({message})")
                    elif excel_com_requested and "Excel COM unavailable" in message:
                        st.warning(f"Excel COM ไม่พร้อมใช้งาน ระบบใช้โหมดสำรองแบบ PNG ({message})")
            if not confirmed_result.workbook_bytes:
                st.error("ส่งออกเสร็จแล้ว แต่ไฟล์ Excel ที่สร้างไม่มีข้อมูล")
            else:
                confirmed_hourly_movement = _hourly_movement_for_ui(confirmed_result, mapping_df)
                chart_pngs = report_chart_pngs(
                    confirmed_hourly_movement,
                    vehicle_composition_report(confirmed_result.normalized),
                    setup=confirmed_setup,
                )
                diagram_png = b""
                if not _is_v2_result(confirmed_result):
                    diagram_png = generate_four_leg_tmc_diagram(
                        confirmed_hourly_movement,
                        confirmed_result.peaks,
                        DiagramConfig(
                            tmc_id=tmc_id,
                            tmc_name=tmc_title,
                            survey_date_text=survey_date_text,
                            north_label=north_label,
                            south_label=south_label,
                            east_label=east_label,
                            west_label=west_label,
                            north_road=north_road,
                            south_road=south_road,
                            east_road=east_road,
                            west_road=west_road,
                            survey_period_text=survey_period,
                            caption_text=caption_text,
                            show_u_turn=show_u_turn,
                        ),
                    )
                st.session_state["tmc_output"] = {
                    "result": confirmed_result,
                    "chart_pngs": dict(chart_pngs),
                    "diagram_png": diagram_png,
                    "workbook_bytes": confirmed_result.workbook_bytes,
                    "workbook_filename": safe_workbook_filename(tmc_id),
                    "confirmed_setup": confirmed_setup,
                    "export_mode": export_mode,
                    "generated_at": export_generated_at,
                }
                set_active_tab("Export")
                _flash_and_rerun("สร้างรายงาน Excel เสร็จแล้ว")
        except Exception as exc:  # pragma: no cover - UI guardrail
            st.error(f"ส่งออกไฟล์ไม่สำเร็จ: {exc}")

    output = st.session_state.get("tmc_output")
    if output:
        _render_section_header("ดาวน์โหลดและรายละเอียด", "ไฟล์ที่สร้างแล้วและชุดประกอบสำหรับตรวจสอบย้อนหลัง")
        _render_download_button("ดาวน์โหลดรายงาน Excel", output["workbook_bytes"], output["workbook_filename"], EXCEL_MIME)
        session_bytes = st.session_state.get("tmc_project_session_bytes")
        session_filename = st.session_state.get("tmc_project_session_filename", "tmc_session.tmcproj.json")
        if not session_bytes:
            session = _build_session_from_state(
                uploaded_file.name if uploaded_file is not None else None,
                len(file_bytes) if uploaded_file is not None else None,
            )
            session_bytes = session_to_json_bytes(session)
            filename_seed = st.session_state.get("tmc_id_input") or st.session_state.get("tmc_title_input") or (uploaded_file.name if uploaded_file is not None else None)
            session_filename = safe_project_session_filename(filename_seed)
            st.session_state["tmc_project_session_bytes"] = session_bytes
            st.session_state["tmc_project_session_filename"] = session_filename

        output_result = output["result"]
        output_setup = output.get("confirmed_setup", setup)
        summary_text = build_export_summary_text(
            setup=output_setup,
            source_file_name=uploaded_file.name if uploaded_file is not None else st.session_state.get("tmc_loaded_source_file_name", ""),
            export_mode=output.get("export_mode", export_mode),
            peaks=output_result.peaks,
            mapping=mapping_df,
            qc=output_result.qc,
            workbook_filename=output["workbook_filename"],
            pce_factors=output_result.pce_factors,
            export_settings={
                "template_version": TEMPLATE_VERSION,
                "template_name": Path(DEFAULT_TEMPLATE_PATH).name,
                "template_map_name": Path(DEFAULT_TEMPLATE_MAP_PATH).name,
            },
            generated_at=output.get("generated_at"),
        )
        mapping_preset_bytes = serialize_mapping_preset(
            build_mapping_preset(
                mapping_df,
                preset_name=str(st.session_state.get("tmc_id_input") or st.session_state.get("tmc_title_input") or "TMC Mapping Preset"),
                movement_code_scheme=_current_mapping_scheme(),
            )
        )
        if _is_v2_result(output_result):
            package_bytes = create_v2_generated_export_package_zip(
                workbook_bytes=output["workbook_bytes"],
                workbook_filename=output["workbook_filename"],
                setup=output_setup,
                peaks=output_result.peaks,
                mapping=mapping_df,
                qc=output_result.qc,
                mapping_preset_bytes=mapping_preset_bytes,
                mapping_preset_filename="mapping_preset.mapping.json",
                source_file_name=uploaded_file.name if uploaded_file is not None else st.session_state.get("tmc_loaded_source_file_name", ""),
                export_mode=output.get("export_mode", export_mode),
                generated_at=output.get("generated_at"),
            )
        else:
            package_bytes = create_export_package_zip(
                workbook_bytes=output["workbook_bytes"],
                workbook_filename=output["workbook_filename"],
                export_summary_text=summary_text,
                project_session_bytes=session_bytes,
                project_session_filename=session_filename,
                mapping_preset_bytes=mapping_preset_bytes,
                mapping_preset_filename="mapping_preset.mapping.json",
                mapping=mapping_df,
                chart_pngs=output.get("chart_pngs", {}),
                diagram_png=output.get("diagram_png"),
            )
        with st.expander("ตัวอย่างไฟล์ใน Export Package ZIP", expanded=False):
            if _is_v2_result(output_result):
                preview_files = [
                    output["workbook_filename"],
                    "export_summary.txt",
                    "mapping_preset.mapping.json",
                    "mapping_table.xlsx",
                    "diagram/movement_diagram_data.csv",
                    "diagram/movement_diagram.png",
                ]
            else:
                preview_files = [
                    output["workbook_filename"],
                    "export_summary.txt",
                    session_filename,
                    "mapping_preset.mapping.json",
                    "mapping_table.xlsx",
                    "charts/hourly_pcu_chart.png",
                    "charts/vehicle_composition_chart.png",
                    "charts/tmc_movement_diagram.png",
                ]
            st.code("\n".join(preview_files), language="text")
            st.caption("Export Package ZIP ไม่รวม raw input Excel")
        _render_download_button(
            "ดาวน์โหลด Export Package ZIP",
            package_bytes,
            safe_package_filename(output["workbook_filename"]),
            PACKAGE_MIME,
        )
        if session_bytes:
            st.download_button(
                "ดาวน์โหลด Project Session",
                data=download_buffer(session_bytes),
                file_name=session_filename,
                mime=PROJECT_SESSION_MIME,
                key="download_project_session_export",
            )
        with st.expander("กราฟและ Diagram สำหรับรายงาน", expanded=False):
            chart_pngs = output["chart_pngs"]
            chart_cols = st.columns(2)
            with chart_cols[0]:
                st.image(chart_pngs["hourly_pcu"], caption="ปริมาณจราจรรวมรายชั่วโมง")
                _render_download_button("ดาวน์โหลดกราฟ PCU รายชั่วโมง (PNG)", chart_pngs["hourly_pcu"], "hourly_pcu_chart.png", PNG_MIME)
            with chart_cols[1]:
                st.image(chart_pngs["vehicle_composition"], caption="สัดส่วนประเภทยานพาหนะ")
                _render_download_button(
                    "ดาวน์โหลดกราฟสัดส่วนยานพาหนะ (PNG)",
                    chart_pngs["vehicle_composition"],
                    "vehicle_composition_chart.png",
                    PNG_MIME,
                )
            if output.get("diagram_png"):
                st.image(output["diagram_png"], caption="Four-leg TMC movement diagram")
                _render_download_button("ดาวน์โหลด Diagram movement (PNG)", output["diagram_png"], "tmc_movement_diagram.png", PNG_MIME)
            elif _is_v2_result(output_result):
                st.caption("approach_movement diagram PNG อยู่ใน Export Package ZIP ที่ path diagram/movement_diagram.png")
    else:
        _render_empty_state(
            "ยังไม่มีไฟล์ส่งออก",
            "กำหนดช่วงเร่งด่วน AM/PM แล้วสร้างรายงาน Excel เมื่อพร้อม",
        )
