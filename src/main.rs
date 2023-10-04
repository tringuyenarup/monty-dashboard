use anyhow::{Ok, Result};
use csv::Reader;
use rust_xlsxwriter::{Format, Workbook, Worksheet};

#[derive(Debug)]
struct Entry {
    group: String,
    metric: String,
    subdivions: String,
    values: Vec<String>,
}

impl Entry {
    fn from_string_record(record: csv::StringRecord) -> Result<Self> {
        let mut line = record.iter();
        let group: String = line.next().unwrap().to_string();
        let metric: String = line.next().unwrap().to_string();
        let subdivions: String = line.next().unwrap().to_string();

        Ok(Entry {
            group,
            metric,
            subdivions,
            values: line.map(|s| s.to_string()).collect::<Vec<String>>(),
        })
    }
}

fn main() -> Result<()> {
    let mut reader = Reader::from_path("inputs/test.csv").unwrap();
    let mut workbook = Workbook::new();

    let scenarios = reader
        .headers()
        .unwrap()
        .iter()
        .skip(3)
        .map(|s| s.to_string())
        .collect::<Vec<String>>();

    let mut worksheet = workbook.add_worksheet();

    let mut previous_subdivisions: Vec<String> = Vec::new();
    let mut previous_metric: String = String::from("");

    let mut row = 1; // reset this with new sheet

    for result in reader.records() {
        let entry = Entry::from_string_record(result?)?;

        if worksheet.name() == *"Sheet1" {
            format_sheet(worksheet, &entry, &scenarios)?;
        } else if worksheet.name() != entry.group {
            worksheet = workbook.add_worksheet();
            format_sheet(worksheet, &entry, &scenarios)?;
            row = 1;
            previous_subdivisions = Vec::new();
            previous_metric = String::from("");
        }
        // Write cell
        worksheet.set_row_height(row, 18)?;
        let mut col = 0;
        let current_cell_subdivisions = entry
            .subdivions
            .split('_')
            .map(String::from)
            .collect::<Vec<String>>();
        // if new metrics
        if previous_metric.is_empty() || previous_metric != entry.metric {
            if row > 1 {
                row += 1;
            }
            write_cell_string(worksheet, &row, &col, &entry.metric, 0)?;
            previous_metric = entry.metric.clone();
            row += 1;
        }
        // now try to read the division or a new sub item
        if previous_subdivisions.is_empty()
            || previous_subdivisions.first().unwrap() != current_cell_subdivisions.first().unwrap()
        {
            // if row > 2 {
            //     row += 1;
            // }
            for (index, val) in current_cell_subdivisions.iter().enumerate() {
                if index != current_cell_subdivisions.len() - 1 {
                    // write string
                    write_cell_string(worksheet, &row, &col, val, (index + 1) as u8)?;
                    row += 1;
                } else {
                    // write leaf
                    write_cell_string(
                        worksheet,
                        &row,
                        &col,
                        val,
                        (current_cell_subdivisions.len() + 1) as u8,
                    )?;
                    col += 1;
                    // write number
                    for value in entry.values.iter() {
                        write_cell_number(worksheet, &row, &col, value)?;
                        col += 1;
                    }
                }
            }
            row += 1;
        } else {
            // find the different subdivision
            if let Some(index) = previous_subdivisions
                .iter()
                .zip(current_cell_subdivisions.iter())
                .position(|(a, b)| a != b)
            {
                // basically write with coresponding indent
                if index != current_cell_subdivisions.len() - 1 {
                    for i in index..current_cell_subdivisions.len() {
                        write_cell_string(
                            worksheet,
                            &row,
                            &col,
                            current_cell_subdivisions.get(index).unwrap(),
                            (i + 1) as u8,
                        )?;
                        row += 1;
                    }
                } else {
                    // write the other cell
                    write_cell_string(
                        worksheet,
                        &row,
                        &col,
                        current_cell_subdivisions.last().unwrap(),
                        (current_cell_subdivisions.len() + 1) as u8,
                    )?;

                    col += 1;
                    for value in entry.values.iter() {
                        write_cell_number(worksheet, &row, &col, value)?;
                        col += 1;
                    }
                    row += 1;
                }
            }
        }
        previous_subdivisions = current_cell_subdivisions;
    }

    workbook.save("outputs/demo_2.xlsx")?;
    Ok(())
}

fn write_cell_string(
    worksheet: &mut Worksheet,
    row: &u32,
    col: &u16,
    data: &str,
    indent: u8,
) -> Result<()> {
    worksheet.write_string_with_format(
        *row,
        *col,
        data,
        &Format::new()
            .set_bold()
            .set_font_size(10)
            .set_font_name("Aptos")
            .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
            .set_align(rust_xlsxwriter::FormatAlign::Left)
            .set_indent(indent),
    )?;

    Ok(())
}

fn write_cell_number(worksheet: &mut Worksheet, row: &u32, col: &u16, value: &str) -> Result<()> {
    worksheet.write_number_with_format(
        *row,
        *col,
        value.parse::<f32>().unwrap(),
        &Format::new()
            .set_num_format("#,##0.00")
            .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
            .set_align(rust_xlsxwriter::FormatAlign::Right)
            .set_indent(1)
            .set_font_size(10)
            .set_font_name("Aptos"),
    )?;
    Ok(())
}

fn format_sheet(worksheet: &mut Worksheet, entry: &Entry, scenarios: &[String]) -> Result<()> {
    worksheet.set_name(entry.group.clone())?;
    // format columns
    worksheet.set_column_width(0, 34.83)?;
    // format row
    worksheet.set_row_height(0, 36)?;
    // write metric name:
    worksheet.write_with_format(
        0,
        0,
        entry.group.clone(),
        &Format::new()
            .set_align(rust_xlsxwriter::FormatAlign::Center)
            .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
            .set_background_color(0xFFD700)
            .set_border_bottom(rust_xlsxwriter::FormatBorder::Thin)
            .set_font_size(14)
            .set_font_name("Aptos"),
    )?;
    let mut col = 1;

    for (index, scenario) in scenarios.iter().enumerate() {
        if index < scenario.len() - 1 {
            worksheet.set_column_width(col, 14.83)?;
            worksheet.write_with_format(
                0,
                col,
                scenario,
                &Format::new()
                    .set_align(rust_xlsxwriter::FormatAlign::Center)
                    .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
                    .set_background_color(0xD3D3D3)
                    .set_border_bottom(rust_xlsxwriter::FormatBorder::Thin)
                    .set_font_size(10)
                    .set_font_name("Aptos"),
            )?;
        } else {
            worksheet.set_column_width(col, 14.83)?;
            worksheet.write_with_format(
                0,
                col,
                scenario,
                &Format::new()
                    .set_align(rust_xlsxwriter::FormatAlign::Center)
                    .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
                    .set_background_color(0xD3D3D3)
                    .set_border_bottom(rust_xlsxwriter::FormatBorder::Thin)
                    .set_border_right(rust_xlsxwriter::FormatBorder::Medium)
                    .set_font_size(10)
                    .set_font_name("Aptos"),
            )?;
        }
        col += 1;
    }

    Ok(())
}
