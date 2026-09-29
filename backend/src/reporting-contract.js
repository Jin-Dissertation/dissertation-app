// Explicit research projections. Never export SELECT * from study tables.
// Raw training details_json includes authentication aliases and is deliberately omitted.
export const REPORTING_DATASETS = [
  {
    name: "aqg_submissions", key: "record_id",
    columns: "record_id participant_id session_id context_id status session_start submitted_at llm_product llm_model llm_description course_context question_context extra_instructions desired_questions final_response feedback_text audio_object_key audio_duration_seconds active_seconds session_close_type session_close_at app_version mode created_at".split(" ")
  },
  {
    name: "aqg_events", key: "record_id",
    columns: "record_id event_id batch_id event_timestamp ms_since_previous participant_id session_id context_id event_type button_id detail_text detail_json device_label created_at".split(" ")
  },
  {
    name: "aqg_feedback", key: "record_id",
    columns: "record_id feedback_id participant_id session_id submitted_at text_feedback audio_object_key audio_original_filename audio_mime_type audio_duration_seconds created_at".split(" ")
  },
  {
    name: "training_submissions", key: "record_id",
    columns: "record_id participant_id session_id session_seq session_start session_end duration_ms duration_formatted active_seconds total_questions correct_first event_count item_count app_version content_version current_card_number device_trail submitted_at created_at".split(" ")
  },
  {
    name: "training_submission_items", key: "record_id",
    columns: "record_id participant_id session_id content_version item_number response_text response_ms created_at".split(" ")
  },
  {
    name: "training_events", key: "record_id",
    columns: "record_id event_id batch_id event_timestamp ms_since_previous participant_id session_id session_seq event_type section_index card_index detail_text detail_json device_label created_at".split(" ")
  },
  {
    name: "training_feedback", key: "record_id",
    columns: "record_id feedback_id participant_id session_id session_seq submitted_at section_index section_title card_index feedback_source text_feedback device_label created_at".split(" ")
  },
  {
    name: "nonparticipant_button_counts", key: "button_id",
    columns: ["button_id", "press_count", "updated_at"]
  }
];

export const REPORTING_PROTOCOL_VERSION = 1;
