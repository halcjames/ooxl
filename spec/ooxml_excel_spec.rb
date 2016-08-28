require 'spec_helper'

describe OOXML::Excel do
  it 'has a version number' do
    expect(OOXML::Excel::VERSION).not_to be nil
  end

  it 'loads spreadsheet' do
    ooxml = OOXML::Excel.new('spec/resources/test.xlsm')
    expect(ooxml.class).to be OOXML::Excel
  end

  # will add on the next update
end
