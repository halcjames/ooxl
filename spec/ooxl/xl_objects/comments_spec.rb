require 'spec_helper'

describe OOXL::Comments do
  let(:comment_sheet_1) { "spec/ooxl/resources/comment_sheet_1.xml" }
  let(:comment_sheet_2) { "spec/ooxl/resources/comment_sheet_2.xml" }

  context 'when parsing comments sheets' do
    it 'correctly parses comments in the format text/t' do
      comments = described_class.load_from_stream(File.read(comment_sheet_1))
      expect(comments.comments.values.all?(&:blank?)).to be_falsey
    end

    it 'correctly parses comments in the format text/r/t' do
      comments = described_class.load_from_stream(File.read(comment_sheet_2))
      expect(comments.comments.values.all?(&:blank?)).to be_falsey
    end
  end
end
