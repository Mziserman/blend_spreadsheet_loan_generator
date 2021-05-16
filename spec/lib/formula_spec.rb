RSpec.describe SpreadsheetLoanGenerator::Formula do
  subject { described_class.new(loan: loan) }

  let(:loan) {
    SpreadsheetLoanGenerator::Loan.new(
      amount: amount,
      duration: duration,
      period_duration: options.fetch(:period_duration),
      rate: rate,
      due_on: options.fetch(:due_on),
      deferred_and_capitalized: options.fetch(:deferred_and_capitalized),
      deferred: options.fetch(:deferred),
      type: options.fetch(:type),
      interests_type: options.fetch(:interests_type),
      starting_capitalized_interests: options.fetch(:starting_capitalized_interests)
    )
  }
  let(:options) do
    {
      type: type,
      due_on: due_on,
      period_duration: period_duration,
      deferred_and_capitalized: deferred_and_capitalized,
      deferred: deferred,
      interests_type: interests_type,
      starting_capitalized_interests: starting_capitalized_interests
    }
  end
  let(:type) { 'standard' }
  let(:due_on) { Date.today }
  let(:period_duration) { 1 }
  let(:deferred_and_capitalized) { 0 }
  let(:deferred) { 0 }
  let(:interests_type) { 'simple' }
  let(:starting_capitalized_interests) { 0.0 }

  let(:line) { rand(2..9999) }

  context '1000 12% 12 months 12/11/2004' do
    let(:rate) { 0.12 }
    let(:amount) { 1000.0 }
    let(:duration) { 12 }

    let(:due_on) { Date.new(2004, 11, 12) }

    it 'returns right indexes' do
      expect(subject.index_formula(line: 2)).to eq(1)

      random_line = line
      expect(subject.index_formula(line: random_line)).to eq(random_line - 1)
    end

    it 'returns right due_on' do
      expect(subject.due_on_formula(line: 2)).to eq(loan.due_on)

      random_line = line
      expect(subject.due_on_formula(line: random_line)).to eq(loan.due_on + (random_line - 2).months)
    end

    it 'returns rounded period_capital' do
      expect(subject.period_capital_formula(line: line)).to match(Regexp.new('=ARRONDI\(.*; ?2\)'))
    end

    it 'returns rounded period_interests' do
      expect(subject.period_interests_formula(line: line)).to match(Regexp.new('=ARRONDI\(.*; ?2\)'))
    end

    it 'returns rounded period_total' do
      expect(subject.period_total_formula(line: line)).to match(Regexp.new('=ARRONDI\(.*; ?2\)'))
    end

    it 'returns standard period_calculated_capital_formula' do
      random_line = line
      expect(subject.period_calculated_capital_formula(line: random_line)).to eq(
        SpreadsheetLoanGenerator::Standard.new(loan: loan)
          .period_calculated_capital_formula(line: random_line)
      )
    end

    it 'returns simple period_rate_formula' do
      random_line = line
      expect(subject.period_rate_formula(line: random_line)).to eq(
        SpreadsheetLoanGenerator::SimpleInterests.new(loan: loan)
          .period_rate_formula(line: random_line)
      )
    end
  end
end
