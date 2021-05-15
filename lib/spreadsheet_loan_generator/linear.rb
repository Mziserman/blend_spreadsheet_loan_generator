module SpreadsheetLoanGenerator
  class Linear
    include SpreadsheetConcern

    attr_accessor :loan
    def initialize(loan:)
      @loan = loan
    end

    def period_calculated_capital_formula(*)
      amount =
        if loan.deferred_and_capitalized.zero?
          excel_float(loan.amount)
        else
          "(#{capitalized_interests_end(loan.deferred_and_capitalized + 1)} + #{excel_float(loan.amount)})"
        end
      "=#{amount} / #{loan.non_deferred_duration}"
    end
  end
end
