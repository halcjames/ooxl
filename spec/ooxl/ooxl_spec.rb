require 'spec_helper'
describe OOXL do
  let(:ooxml) { OOXL.new('spec/ooxl/resources/test.xlsx', padded_rows: true) }

  it 'loads spreadsheet' do
    expect(ooxml.class).to be OOXL
  end

  it 'loads sheets' do
    expect(ooxml.sheets).to eq ['Sheet1', 'Sheet2', 'Hidden']
    expect(ooxml.sheets(skip_hidden: true)).to eq ['Sheet1', 'Sheet2']
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

  it 'loads list values' do
    expect(ooxml.list_values('"Demo,Demo 2"')).to eq ['Demo', 'Demo 2']
    expect(ooxml.list_values('Sheet2!A1:B2')). to eq [['Range Value', '2'], ['Range Value 2', '3']]
    expect(ooxml.list_values('Sheet2!A1:A2')). to eq ['Range Value', 'Range Value 2']
  end

  it 'loads padded rows' do
    values = []
    ooxml.sheet('Sheet2').each do |row|
      if row.cells.blank?
        values << []
      else
        values << row.cells.map(&:value)
      end
    end
    expect(values.size).to eq 16
    expect(values.last).to eq ["Very Far", "5"]
  end

  # will add on the next update
end
