module SpreadsheetLoanGenerator
  class Linear
    include SpreadsheetConcern

    attr_accessor :loan
    def initialize(loan:)
      @loan = loan
    end

    def standard_params(line:)
      if loan.deferred_and_capitalized.zero?
        amount = excel_float(loan.amount)
        term_cell = "#{column_letter[:index]}#{line}"
      else
        amount = "#{capitalized_interests_end(loan.deferred_and_capitalized + 1)} + #{excel_float(loan.amount)}"
        term_cell = "#{column_letter[:index]}#{line} - #{loan.total_deferred_duration}"
      end

      "#{period_rate(line)};#{term_cell};#{loan.non_deferred_duration};#{amount}"
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

    def period_interests_formula(line:)
      "=ARRONDI(#{period_calculated_interests(line)}; 2)"
    end
  end
end
