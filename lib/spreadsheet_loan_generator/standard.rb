module SpreadsheetLoanGenerator
  class Standard
    include SpreadsheetConcern

    attr_accessor :loan

    def initialize(loan:)
      @loan = loan
    end

    def standard_params(line:)
      amount =
        if loan.deferred_and_capitalized.zero?
          excel_float(loan.amount)
        else
          "#{capitalized_interests_end(loan.deferred_and_capitalized + 1)} + #{excel_float(loan.amount)}"
        end
      term_cell = "#{column_letter[:index]}#{line} - #{loan.total_deferred_duration}"

      "#{period_rate(line)};#{term_cell};#{loan.non_deferred_duration};#{amount}"
    end

    def period_calculated_capital_formula(line:)
      "=-PPMT(#{standard_params(line: line)})"
    end

    def period_calculated_interests_formula(line:)
      "=-IPMT(#{standard_params(line: line)})"
    end
  end
end

