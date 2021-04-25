//OB VSAKI SPREMEMBI KODE JE TREBA NA NOVO DEPLOYATI
//IN KOPIRATI NOV LINK V WOOCOMMERCE WEBHOOK

//this is a function that fires when the webapp receives a GET request
function doGet(e) {
  return HtmlService.createHtmlOutput("success!"); 
}

//this is a function that fires when the webapp receives a POST request
function doPost(e) {
  let myData = JSON.parse([e.postData.contents]);

  //
  if (!myData || myData === undefined) return;
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  let activeSheet = sheet.getActiveSheet();
  let dataRange = activeSheet.getDataRange();
  let values = dataRange.getValues();

  let final = [];

  let line_items = myData.line_items || ""; // the || "" at the end is the safeguard for undefined values
  let meta_data = myData.meta_data || "";
  let order_id = myData.id || "";
  let order_created = myData.date_created || "";
  let order_status = myData.status || "";
  let billing_email = myData.billing.email || "";
  let billing_first_name = myData.billing.first_name || "";
  let billing_last_name = myData.billing.last_name || "";
  let billing_company = myData.billing.company || "";
  let order_total = Number(myData.total) || "";
  let order_total_tax = Number(myData.total_tax) || "";
  let order_shipping = Number(myData.shipping_total) || "";
  let order_shipping_tax = Number(myData.shipping_tax) || "";
  let total = order_total + order_total_tax + order_shipping + order_shipping_tax || "";
  let discount_amount = "" || "";
  let discount_amount_tax = "" || "";
  let coupon_code = "" || "";
  let payment_method_title = myData.payment_method_title || "";
  let shipping_method_title = myData.shipping_lines.method_title || "";


  let utm_medium, utm_campaign, utm_source_first, utm_source, utm_sess_ref,
    utm_sess_ref_clean, utm_sess_landing, utm_sess_landing_clean,
    gclid_visit, gclid_visit_date, fbclid_url,
    fbclid_visit, fbclid_visit_date, gclid_url,
    gclid_url_clean, fbclid_url_clean, conversion_type, conversion_date = "";

  for (let i = 0; i < meta_data.length; i++) {
    const meta_element = meta_data[i];
    //shorthand operatorji, ki pogledajo če je meta_key enak določeni vrednosti
    //Če je pogoj true določijo meta_value ustrezni spremenljivki
    // del kode: || "" je safeguard za undefined vrednost 
    (meta_element.key == "_afl_wc_utm_utm_medium_1st") ? utm_medium = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_utm_campaign_1st") ? utm_campaign = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_utm_source") ? utm_source = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_sess_referer") ? utm_sess_ref = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_sess_referer_clean") ? utm_sess_ref_clean = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_sess_landing") ? utm_sess_landing = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_sess_landing_clean") ? utm_sess_landing_clean = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_gclid_visit") ? gclid_visit = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_gclid_visit_date_local") ? gclid_visit_date = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_fbclid_url") ? fbclid_url = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_fbclid_url_clean") ? fbclid_url_clean = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_fbclid_visit") ? fbclid_visit = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_utm_source_1st") ? utm_source_first = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_conversion_type") ? conversion_type = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_conversion_date_local") ? conversion_date = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_fbclid_visit_date_local") ? fbclid_visit_date = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_gclid_url") ? gclid_url = meta_element.value || "" : null;
    (meta_element.key == "_afl_wc_utm_gclid_url_clean") ? gclid_url_clean = meta_element.value || "" : null;
  }

  for (let i = 0; i < line_items.length; i++) {
    let line_item = line_items[i] || "";
    let line_id = line_item.id || "";
    let product_name = line_item.name || "";
    let product_id = line_item.product_id || "";
    let product_qty = line_item.quantity || "";
    let product_total = line_item.total || "";
    let product_total_tax = line_item.total_tax || "";
    let item_cost = product_total + product_total_tax || "";
    let sku = line_item.sku || "";
    let line_meta = line_item.meta_data || "";
    let upsell = "" || "";
    let item_number = i + 1 || "";
    let product_subtotal = Number(line_item.subtotal) || "";
    let product_subtotal_tax = Number(line_item.subtotal_tax) || "";
    let line_subtotal = product_subtotal + product_subtotal_tax || "";

    let line_id_column_number = 40; //spredsheet column AO (basically 41 substracted by 1 as index starts with 0)
    let found_existing_row = 0;
    //find if line_item id exists to replace it or not
    for (let k = 0; k < values.length; k++) {
      if (values[k][line_id_column_number] == line_id) {
        found_existing_row = k + 1;
        break;
      }
    }

    //loop through line items' meta data
    for (let j = 0; j < line_meta.length; j++) {
      let meta_item = line_meta[j];

      (meta_item.key == "_upstroke_purchase") ? upsell = meta_item.value: null;
    }

    final = [
      order_id,
      order_status,
      order_created.replace("T", " "), // need to modify the timestamp format
      utm_medium,
      utm_campaign,
      utm_source_first,
      utm_source,
      utm_sess_ref,
      utm_sess_ref_clean,
      utm_sess_landing,
      utm_sess_landing_clean,
      gclid_visit,
      gclid_visit_date,
      fbclid_url,
      fbclid_visit,
      fbclid_visit_date,
      gclid_url_clean,
      gclid_url,
      fbclid_url_clean,
      conversion_type,
      conversion_date,
      item_number,
      upsell,
      product_name,
      product_qty,
      product_total,
      sku,
      billing_first_name,
      billing_last_name,
      billing_company,
      order_total,
      total,
      order_total_tax,
      coupon_code,
      discount_amount,
      discount_amount_tax,
      payment_method_title,
      shipping_method_title,
      product_subtotal_tax,
      line_subtotal,
      line_id
    ]

    if (found_existing_row != 0) {
      let result = sheet.getRange("A" + found_existing_row + ":AO" + found_existing_row).setValues([final]); //mofo setvalues prejema 2D array
    } else {
      sheet.appendRow(final);
    }
  }
}
