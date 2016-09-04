require 'spec_helper'
require 'pry'
describe OOXML::Excel::Sheet do
  let(:ooxml) { OOXML::Excel.new('spec/ooxml_excel/resources/test.xlsx') }

  it 'loads sheet' do
    expect(ooxml.sheet('Sheet1').class).to be OOXML::Excel::Sheet
    expect(ooxml['Sheet1'].class).to be OOXML::Excel::Sheet
  end

  it 'loads font object' do
    font = ooxml.sheet('Sheet1').font('A1')
    expect(font.bold?).to be false
    expect(font.name).to eq "Arial"
    expect(font.size).to eq "10"
    expect(font.rgb_color).to eq "FFFF3333"
  end

  it 'loads fill object' do
    fill = ooxml.sheet('Sheet1').fill('A2')

    expect(fill.pattern_type).to eq 'solid'
    expect(fill.bg_color).to eq 'FFFF6600'
    expect(fill.fg_color).to eq 'FFFF3333'
  end

  it 'loads data validations' do
    data_validation = ooxml.sheet('Sheet1').data_validation('B1')
    expect(data_validation.allow_blank).to eq "true"
    expect(data_validation.formula).to eq "0"
    expect(data_validation.prompt).to eq "Test"
    expect(data_validation.type).to eq "none"
  end

  it 'loads column' do
    column = ooxml.sheet('Sheet1').column('C')
    expect(column.class). to be OOXML::Excel::Sheet::Column
    expect(column.hidden?).to be false
    expect(column.width).to eq '0'
    expect(column.id).to eq '3'
  end

  it 'loads row' do
    row = ooxml.sheet('Sheet1').row(1)
    expect(row.class). to be OOXML::Excel::Sheet::Row
  end
end
