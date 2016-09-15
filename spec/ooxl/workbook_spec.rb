require 'spec_helper'
require 'pry'
describe OOXL::Workbook do
  let(:workbook_xml) do
    '<workbook mc:Ignorable="x15">
      <sheets>
        <sheet name="Sheet1" sheetId="1" id="rId1"/>
        <sheet name="Sheet2" sheetId="2" id="rId2"/>
      </sheets>
      <definedNames>
        <definedName name="named_range">Sheet2!$A$1:$A$2</definedName>
        <definedName name="range">Sheet2!$A$1:$B$2</definedName></definedNames>
     </workbook>
    '
  end

  let(:workbook) { OOXL::Workbook.load_from_stream(workbook_xml) }

  it 'loads workbook' do
    expect(workbook.class).to be OOXL::Workbook
  end

  it 'loads sheets' do
    expect(workbook.sheets.size).to eq 2
    expect(workbook.sheets[0][:name]).to eq 'Sheet1'
  end

  it 'loads defined names' do
    first_key = workbook.defined_names.keys.first
    expect(workbook.defined_names.size).to eq 2
    expect(workbook.defined_names[first_key]).to eq 'Sheet2!$A$1:$A$2'
  end
end
