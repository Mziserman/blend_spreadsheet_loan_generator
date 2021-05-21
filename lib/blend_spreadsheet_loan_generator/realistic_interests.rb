module BlendSpreadsheetLoanGenerator
  class RealisticInterests
    include SpreadsheetConcern

    attr_accessor :loan
    def initialize(loan:)
      @loan = loan
    end

    def period_rate_formula(line:)
      "=#{excel_float(loan.rate)} * ((#{period_leap_days(line)} / 366) + (#{period_non_leap_days(line)} / 365))"
    end

    def period_fees_rate_formula(line:)
      "=#{excel_float(loan.fees_rate)} * ((#{period_leap_days(line)} / 366) + (#{period_non_leap_days(line)} / 365))"
    end
  end
end
