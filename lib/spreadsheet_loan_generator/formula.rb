module SpreadsheetLoanGenerator
  class Formula
    include SpreadsheetConcern

    attr_accessor :loan

    def initialize(loan:)
      @loan = loan
      @interests_formula = loan.interests_formula
      @loan_type_formula = loan.lona_type_formula
    end

    def index_formula(line:)
      line - 1
    end

    def due_on_formula(line:)
      term = line - 1
      loan.due_on + (term * loan.period_duration).months
    end

    def remaining_capital_start_formula(line:)
      return excel_float(loan.amount) if line == 2

      "=#{excel_float(loan.amount)} - #{total_paid_capital_end_of_period(line - 1)}"
    end

    def remaining_capital_end_formula(line:)
      "=#{excel_float(loan.amount)} - #{total_paid_capital_end_of_period(line)}"
    end

    def period_calculated_interests_formula(line:)
      amount_to_capitalize = [
        '(',
        remaining_capital_start(line),
        '+',
        capitalized_interests_start(line),
        ')'
      ].join(' ')
      "=ARRONDI(#{amount_to_capitalize} * #{period_rate(line)}; 14)"
    end

    def accrued_delta_formula(line:)
      return excel_float(0.0) if line == 2

      "=ARRONDI(#{accrued_delta(line - 1)} + #{delta(line)}; 14)"
    end

    def total_paid_capital_end_of_period_formula(line:)
      return "=#{period_capital(line)}" if line == 2

      "=ARRONDI(SOMME(#{column_range(column: period_capital, upto: line)}); 2)"
    end

    def total_paid_interests_end_of_period_formula(line:)
      return "=#{period_interests(line)}" if line == 2

      "=ARRONDI(SOMME(#{column_range(column: period_interests, upto: line)}); 2)"
    end

    def period_capital_formula(line:)
      if line == loan.duration + 1
        "=ARRONDI(#{excel_float(loan.amount)} - #{total_paid_capital_end_of_period(line - 1)}; 2)"
      else
        "=ARRONDI(#{period_calculated_capital(line)} - #{period_reimbursed_capitalized_interests(line)}; 2)"
      end
    end

    def period_total_formula(line:)
      total = [
        period_capital(line),
        '+',
        period_interests(line),
        '+',
        period_reimbursed_capitalized_interests(line)
      ].join(' ')
      "=ARRONDI(#{total}; 2)"
    end

    def capitalized_interests_start_formula(line:)
      return excel_float(0.0) if line == 2

      "=ARRONDI(#{capitalized_interests_end(line - 1)}; 2)"
    end

    def delta_formula(line:)
      amount_added =
        if line == 2
          excel_float(0.0)
        elsif line > 2
          "SOMME(#{column_range(column: amount_to_add, upto: line - 1)})"
        end
      "=ARRONDI(#{period_calculated_interests(line)} - #{period_interests(line)} - (#{capitalized_interests_end(line)} - #{capitalized_interests_start(line)}) - #{period_reimbursed_capitalized_interests(line)}; 14) - #{amount_added}"
    end

    def amount_to_add_formula(line:)
      "=TRONQUE(#{accrued_delta(line)}; 2)"
    end

    def period_reimbursed_capitalized_interests_formula(line:)
      term = line - 1
      if term <= loan.deferred_and_capitalized + loan.deferred
        excel_float(0.0)
      else
        "=ARRONDI(MIN(#{period_calculated_capital(line)}; #{capitalized_interests_start(line)}); 2)"
      end
    end

    def capitalized_interests_end_formula(line:)
      term = line - 1
      if term <= loan.deferred_and_capitalized
        "=ARRONDI(#{capitalized_interests_start(line)} + #{period_calculated_interests(line)}; 2)"
      else
        "=ARRONDI(#{capitalized_interests_start(line)} - #{period_reimbursed_capitalized_interests(line)}; 2)"
      end
    end

    def period_calculated_capital_formula(line:)
      term = line - 1
      if term <= loan.deferred_and_capitalized
        excel_float(0.0)
      elsif term <= loan.deferred_and_capitalized + loan.deferred
        excel_float(0.0)
      else
        @loan_type_formula.period_calculated_capital_formula(line: line)
      end
    end

    def period_interests_formula(line:)
      term = line - 1
      if term <= loan.deferred_and_capitalized
        excel_float(0.0)
      elsif term <= loan.deferred_and_capitalized + loan.deferred
        "=ARRONDI(#{period_calculated_interests(line)}; 2)"
      else
        @loan_type_formula.period_interests_formula(line: line)
      end
    end

    def period_rate_formula(line:)
      @interests_formula.period_rate_formula(line: line)
    end
  end
end
