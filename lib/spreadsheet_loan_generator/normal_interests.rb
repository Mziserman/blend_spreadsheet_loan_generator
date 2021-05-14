module SpreadsheetLoanGenerator
  class NormalInterests
    include SpreadsheetConcern

    attr_accessor :loan
    def initialize(loan:)
      @loan = loan
    end

    def period_rate_formula(*)
      "=TAUX.NOMINAL(#{excel_float(loan.rate)};#{excel_float(12.0 / loan.period_duration)})"
    end
  end
end
