// QRE Analytic Reframe v2 — Google Apps Script
// Paste this entire file into Extensions → Apps Script in your Google Sheet.
// Then: Deploy → New Deployment → Web App
//   Execute as: Me
//   Who has access: Anyone
// Copy the deployment URL and paste it into index.html as SUBMIT_URL.

const SHEET_NAME = 'Responses';

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    const d = JSON.parse(e.postData.contents);

    if (sheet.getLastRow() === 0) {
      const hdr = sheet.appendRow(getHeaders());
      sheet.getRange(1, 1, 1, getHeaders().length)
        .setFontWeight('bold')
        .setBackground('#1e293b')
        .setFontColor('white');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow(buildRow(d));

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── HEADERS ──────────────────────────────────────────────────────────────────
function getHeaders() {
  const elCols = [1, 2, 3, 4].flatMap(n => [
    `q1_el${n}_description`, `q1_el${n}_timestamp`, `q1_el${n}_classification`,
    `q1_el${n}_humorous`, `q1_el${n}_severity`, `q1_el${n}_pivotal`
  ]);
  return [
    'submitted_at', 'ad_id', 'coder_name', 'date', 'q1_element_count',
    ...elCols,
    'q4_dramatic_tension', 'q4_evidence',
    'q5_character_structure', 'q5_evidence',
    'q5a_protagonist_clarity', 'q5a_evidence',
    'q6_protagonist_goal', 'q6_evidence',
    'q7_stakes', 'q7_evidence',
    'q8_obstacle_score', 'q8_antagonist_present', 'q8_evidence',
    'q9_viewer_care', 'q9_evidence',
    'q10_peak_moment', 'q10_evidence',
    'q10a_emotional_peak', 'q10a_evidence',
    'q11_resolution_type', 'q11_evidence',
    'q12_resolution_locus', 'q12_evidence',
    'q13_tone', 'q13_evidence',
    'q14_brand_role', 'q14_evidence',
    'q15_brand_narrative_position', 'q15_evidence',
    'q16_brand_swap_a', 'q16_evidence',
    'q17_brand_swap_b', 'q17_derived_d2_tier', 'q17_evidence',
    'q18_brand_dimension', 'q18_evidence',
    'q19_communication_mode', 'q19_evidence',
    'q20_verbal_speaker',
    'q21_cultural_elements', 'q21_evidence',
    'q22_prior_knowledge', 'q22_evidence',
    'q23_cultural_ref_type', 'q23_evidence',
    'q24_cultural_ref_load', 'q24_evidence',
    'q25_visual_unusualness', 'q25_evidence',
    'q26_visual_category_diff', 'q26_evidence',
    'q28_first_5_seconds', 'q28_evidence',
    'q29_message_clarity', 'q29_evidence',
    'q30_audience_insight', 'q30_evidence',
    'q30a_target_specificity', 'q30a_evidence',
    'q31_in_group', 'q31_evidence',
    'q32_out_group', 'q32_evidence',
    'q33_group_breadth', 'q33_evidence',
    'q34_group_framing', 'q34_evidence',
    'q35_interpretive_work', 'q35_evidence',
    'q35a_brand_thinking_moment', 'q35a_evidence',
    'q36_brand_differentiators', 'q36_evidence',
    'q37_benefit_clarity', 'q37_evidence',
    'q38_benefit_type', 'q38_evidence',
    'q39_benefit_framing', 'q39_evidence',
    'q40_cta', 'q40_evidence',
    'q41_pov', 'q41_evidence',
    'q42_pov_boldness', 'q42_evidence',
    'q43_pov_centrality', 'q43_evidence',
    'q44_lifestyle_cues', 'q44_evidence',
    'q45_lifestyle_description',
    'q46_lifestyle_group_type', 'q46_evidence',
    'q47_minority_audience', 'q47_evidence',
    'q48_minority_accessibility',
    'q49_dominant_strategy', 'q49_evidence',
    'q50_brand_product_focus', 'q50_evidence',
    'confidence_scores',
  ];
}

// ─── ROW BUILDER ──────────────────────────────────────────────────────────────
function buildRow(d) {
  const v = x => (x == null ? '' : x);
  const a = x => Array.isArray(x) ? x.join(' | ') : v(x);
  const el = i => d.q1_unexpected_elements[i] || {};

  return [
    new Date().toISOString(),
    v(d.metadata.ad_id),
    v(d.metadata.coder_name),
    v(d.metadata.date),
    v(d.q1_element_count),

    // Up to 4 elements
    ...[0, 1, 2, 3].flatMap(i => [
      v(el(i).description), v(el(i).timestamp), a(el(i).classification),
      v(el(i).humorous), v(el(i).severity), v(el(i).pivotal)
    ]),

    a(d.q4_dramatic_tension.mechanisms),        v(d.q4_dramatic_tension.evidence),
    v(d.q5_character_structure.value),          v(d.q5_character_structure.evidence),
    v(d.q5a_protagonist_clarity.value),         v(d.q5a_protagonist_clarity.evidence),
    v(d.q6_protagonist_goal.value),             v(d.q6_protagonist_goal.evidence),
    a(d.q7_stakes.values),                      v(d.q7_stakes.evidence),
    v(d.q8_obstacle.score),                     v(d.q8_obstacle.antagonist_present), v(d.q8_obstacle.evidence),
    v(d.q9_viewer_care.score),                  v(d.q9_viewer_care.evidence),
    v(d.q10_peak_moment.score),                 v(d.q10_peak_moment.evidence),
    v(d.q10a_emotional_peak.score),             v(d.q10a_emotional_peak.evidence),
    v(d.q11_resolution_type.value),             v(d.q11_resolution_type.evidence),
    v(d.q12_resolution_locus.value),            v(d.q12_resolution_locus.evidence),
    v(d.q13_tone.value === 'Other' ? d.q13_tone.other : d.q13_tone.value), v(d.q13_tone.evidence),
    v(d.q14_brand_role.value),                  v(d.q14_brand_role.evidence),
    a(d.q15_brand_narrative_position.values),   v(d.q15_brand_narrative_position.evidence),
    v(d.q16_brand_swap_a.value),                v(d.q16_brand_swap_a.evidence),
    v(d.q17_brand_swap_b.value),                v(d.q17_brand_swap_b.derived_d2_tier), v(d.q17_brand_swap_b.evidence),
    a(d.q18_brand_dimension.values),            v(d.q18_brand_dimension.evidence),
    v(d.q19_communication_mode.value),          v(d.q19_communication_mode.evidence),
    v(d.q20_verbal_speaker.value),
    a(d.q21_cultural_elements.values),          v(d.q21_cultural_elements.evidence),
    v(d.q22_prior_knowledge_requirement.value), v(d.q22_prior_knowledge_requirement.evidence),
    a(d.q23_cultural_reference_type.values),    v(d.q23_cultural_reference_type.evidence),
    v(d.q24_cultural_reference_load.score),     v(d.q24_cultural_reference_load.evidence),
    v(d.q25_visual_unusualness.score),          v(d.q25_visual_unusualness.evidence),
    v(d.q26_visual_category_difference.score),  v(d.q26_visual_category_difference.evidence),
    a(d.q28_first_5_seconds.values),            v(d.q28_first_5_seconds.evidence),
    v(d.q29_message_clarity.value),             v(d.q29_message_clarity.evidence),
    v(d.q30_audience_insight.score),            v(d.q30_audience_insight.evidence),
    v(d.q30a_target_specificity.score),         v(d.q30a_target_specificity.evidence),
    v(d.q31_in_group.score),                    v(d.q31_in_group.evidence),
    v(d.q32_out_group.score),                   v(d.q32_out_group.evidence),
    v(d.q33_group_breadth.value),               v(d.q33_group_breadth.evidence),
    v(d.q34_group_framing.value),               v(d.q34_group_framing.evidence),
    v(d.q35_interpretive_work.score),           v(d.q35_interpretive_work.evidence),
    v(d.q35a_brand_thinking_moment.value),      v(d.q35a_brand_thinking_moment.evidence),
    a(d.q36_brand_differentiators.values),      v(d.q36_brand_differentiators.evidence),
    v(d.q37_benefit_clarity.score),             v(d.q37_benefit_clarity.evidence),
    a(d.q38_benefit_type.values),               v(d.q38_benefit_type.evidence),
    v(d.q39_benefit_framing.value),             v(d.q39_benefit_framing.evidence),
    v(d.q40_cta.value),                         v(d.q40_cta.evidence),
    v(d.q41_pov.value),                         v(d.q41_pov.evidence),
    v(d.q42_pov_boldness.value),                v(d.q42_pov_boldness.evidence),
    v(d.q43_pov_centrality.score),              v(d.q43_pov_centrality.evidence),
    v(d.q44_lifestyle_cues.value),              v(d.q44_lifestyle_cues.evidence),
    v(d.q45_lifestyle_description),
    a(d.q46_lifestyle_group_type.values),       v(d.q46_lifestyle_group_type.evidence),
    v(d.q47_minority_audience.value),           v(d.q47_minority_audience.evidence),
    v(d.q48_minority_accessibility.value),
    v(d.q49_dominant_strategy.value),           v(d.q49_dominant_strategy.evidence),
    v(d.q50_brand_product_focus.value),         v(d.q50_brand_product_focus.evidence),
    JSON.stringify(d.confidence_scores || {}),
  ];
}
