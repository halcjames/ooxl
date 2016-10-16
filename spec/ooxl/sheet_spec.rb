require 'spec_helper'
require 'pry'
describe OOXL::Sheet do
  let(:sheet_xml) do
    '<worksheet mc:Ignorable="x14ac">
      <cols>
        <col min="1" max="2" width="8.5703125"/>
        <col min="3" max="3" width="0" hidden="1"/>
        <col min="4" max="1025" width="8.5703125"/>
      </cols>
      <sheetData>
      <row r="1" spans="1:2" x14ac:dyDescent="0.2">
        <c r="A1" s="1" t="s"></c>
        <c r="B1" s="2" t="s"><v>1</v></c>
        <c r="C1" s="2" t="s"><v>2</v></c>
      </row>
      <row r="2" spans="1:2" x14ac:dyDescent="0.2">
        <c r="A2" s="3" t="s"><v>2</v></c>
        <c r="B2" s="2" t="s"><v>3</v></c>
        <c r="C3" s="2" t="s"><v>4</v></c>
      </row>
      <mergeCells count="1">
        <mergeCell ref="C1:C3"/>
      </mergeCells>
      <dataValidations count="1">
        <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B6">
          <formula1>named_range</formula1>
        </dataValidation>
      </dataValidations>
    </worksheet>'
  end

  let(:sheet) { OOXL::Sheet.new(sheet_xml, []) }

  it 'loads sheet' do
    expect(sheet.class).to be OOXL::Sheet
  end

  it 'loads columns' do
    expect(sheet.columns.size).to eq 3
    expect(sheet.column('A').class).to be OOXL::Column
    expect(sheet.column('B').class).to be OOXL::Column
    expect(sheet.column('C').class).to be OOXL::Column
  end

  it 'loads row' do
    expect(sheet.rows.size).to eq 2
    expect(sheet.row(1).class).to be OOXL::Row
    expect(sheet.row(1).id).to eq '1'
    expect(sheet.row(2).class).to be OOXL::Row
    expect(sheet.row(2).id).to eq '2'
  end

  it 'loads cells by column' do
    columns = sheet.cells_by_column('A')
    expect(columns.first.class).to be OOXL::Cell
    expect(columns.size).to eq 2
  end

  it 'detects merged cell' do
    expect(sheet.in_merged_cells?('C1')).to be true
    expect(sheet.in_merged_cells?('B1')).to be false
  end

  it 'loads a single cell' do
    expect(sheet.cell('A1').class).to be OOXL::Cell
  end

  it 'loads cell coordinates' do
    cell = sheet.cell('A1')
    expect(cell.column).to eq 'A'
    expect(cell.row).to eq '1'
  end

  it 'loads data validations' do
    expect(sheet.data_validations.size).to eq 1
    expect(sheet.data_validation('B6').class).to eq OOXL::Sheet::DataValidation
  end
 end
