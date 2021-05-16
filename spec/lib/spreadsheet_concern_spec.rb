RSpec.describe SpreadsheetLoanGenerator::SpreadsheetConcern do
  klass = Class.new do
    include SpreadsheetLoanGenerator::SpreadsheetConcern
  end.new

  columns = klass.columns

  subject { klass }
  let(:line) { rand(1..9999) }
  let(:alphabet) { ('A'..'ZZ').to_a }

  columns.each do |column|
    describe "##{column}" do
      it "responds to #{column}" do
        expect(subject.respond_to?(column)).to be true
      end

      it 'returns just the letter without an argument' do
        expect(subject.send(column)).to eq(alphabet[columns.index(column)])
      end

      it 'returns letter/argument couple with an argument' do
        expect(subject.send(column, line)).to eq(alphabet[columns.index(column)] + line.to_s)
      end
    end
  end
end
