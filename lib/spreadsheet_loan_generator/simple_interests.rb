module SpreadsheetLoanGenerator
  class SimpleInterests
    include SpreadsheetConcern

    attr_accessor :loan
    def initialize(loan:)
      @loan = loan
    end

    def period_rate_formula(*)
      "=#{excel_float(loan.rate)} * #{loan.period_duration} / 12,0"
    end
  end
end
