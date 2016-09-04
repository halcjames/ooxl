require 'spec_helper'
require 'pry'
describe OOXML::Excel do
  let(:ooxml) { OOXML::Excel.new('spec/ooxml_excel/resources/test.xlsx') }

  it 'loads spreadsheet' do
    expect(ooxml.class).to be OOXML::Excel
  end

  it 'loads all sheets' do
    expect(ooxml.sheets).to eq ['Sheet1', 'Sheet2']
  end

  it 'loads named range values' do
    expect(ooxml.named_range('named_range')).to eq ['Range Value', 'Range Value 2']
  end

  it 'loads cell range values' do
    expect(ooxml['Sheet2!A1:A2']).to eq ['Range Value', 'Range Value 2']
  end

  it 'loads cell range values (box type)' do
    expect(ooxml['Sheet2!A1:B2']).to eq [['Range Value', '2'], ['Range Value 2', '3']]
  end


  # will add on the next update
end
