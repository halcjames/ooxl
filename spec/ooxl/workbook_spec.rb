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

  let(:relationship_xml) do 
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships 
        xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
        <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
      </Relationships>
    '
  end 

  let(:workbook) { OOXL::Workbook.load_from_stream(workbook_xml, OOXL::Relationships.new(relationship_xml)) }

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
