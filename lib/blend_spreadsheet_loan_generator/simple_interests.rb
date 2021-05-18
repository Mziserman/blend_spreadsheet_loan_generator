module BlendSpreadsheetLoanGenerator
  class SimpleInterests
    include SpreadsheetConcern

    attr_accessor :loan
    def initialize(loan:)
      @loan = loan
    end

    def period_rate_formula(*)
      "=#{excel_float(loan.rate)} * #{loan.period_duration} / 12,0"
    end

    def period_fees_rate_formula(*)
      "=#{excel_float(loan.fees_rate)} * #{loan.period_duration} / 12,0"
    end
  end
end
