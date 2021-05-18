module BlendSpreadsheetLoanGenerator
  class NormalInterests
    include SpreadsheetConcern

    attr_accessor :loan
    def initialize(loan:)
      @loan = loan
    end

    def period_rate_formula(*)
      periods_per_year = excel_float(12.0 / loan.period_duration)
      "=TAUX.NOMINAL(#{excel_float(loan.rate)};#{periods_per_year}) / #{periods_per_year}"
    end


    def period_fees_rate_formula(*)
      periods_per_year = excel_float(12.0 / loan.period_duration)
      "=TAUX.NOMINAL(#{excel_float(loan.fees_rate)};#{periods_per_year}) / #{periods_per_year}"
    end
  end
end
