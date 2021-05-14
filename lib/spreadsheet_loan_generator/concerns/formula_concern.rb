module SpreadsheetLoanGenerator
  module FormulaConcern
    extend ActiveSupport::Concern

    included do
      # use method missing defined in spreadsheet concerns a lot

      def remaining_capital_start_formula(line:)
        base = excel_float(float: loan.amount)

        return base if line == 2

        "=#{base} - #{total_paid_capital_end_of_period(line - 1)}"
      end

      def remaining_capital_end_formula(line:)
        base = excel_float(float: loan.amount)

        "=#{base} - #{total_paid_capital_end_of_period(line)}"
      end

      def period_calculated_interests_formula(line:)
        "=ARRONDI((#{remaining_capital_start(line)} + #{capitalized_interests_start(line)}) * #{period_rate(line)}; 14)"
      end

      def period_accrued_delta_formula(line:)
        return excel_float(float: 0.0) if line == 2

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

      def period_rate_formula
        case loan.interests_type
        when 'normal'
          "=TAUX.NOMINAL(#{excel_float(float: loan.rate)};#{excel_float(float: 12.0 / loan.period_duration)})"
        when 'simple'
          "=#{excel_float(float: loan.rate)} * #{loan.period_duration} / 12,0"
        end
      end

      def period_interests_formula(line:)
        term = line - 1
        if term <= loan.deferred_and_capitalized
          excel_float(float: 0.0)
        elsif term <= loan.deferred_and_capitalized + loan.deferred
          "=ARRONDI(#{period_calculated_interests(line)}; 2)"
        else
          "=ARRONDI(-IPMT(#{standard_params(line: line)}); 2)"
        end
      end

      def period_capital_formula(line:)
        if line == loan.duration + 1
          "=ARRONDI(#{excel_float(float: loan.amount)} - #{total_paid_capital_end_of_period(line - 1)}; 2)"
        else
          "=ARRONDI(#{period_calculated_capital(line)} - #{period_reimbursed_capitalized_interests(line)}; 2)"
        end
      end

      def standard_params(line:)
        if loan.deferred_and_capitalized.zero?
          amount = excel_float(float: loan.amount)
          term_cell = "#{column_letter[:index]}#{line}"
        else
          amount = "#{capitalized_interests_end(loan.deferred_and_capitalized + 1)} + #{excel_float(float: loan.amount)}"
          term_cell = "#{column_letter[:index]}#{line} - #{loan.total_deferred_duration}"
        end

        "#{period_rate(line)};#{term_cell};#{loan.non_deferred_duration};#{amount}"
      end

      def period_calculated_capital_formula(line:)
        term = line - 1
        if term <= loan.deferred_and_capitalized
          excel_float(float: 0.0)
        elsif term <= loan.deferred_and_capitalized + loan.deferred
          excel_float(float: 0.0)
        else
          "=-PPMT(#{standard_params(line: line)})"
        end
      end

      def period_total_formula(line:)
        "=ARRONDI(#{period_capital(line)} + #{period_interests(line)}; 2)"
      end

      def capitalized_interests_start_formula(line:)
        return excel_float(float: 0.0) if line == 2

        "=ARRONDI(#{capitalized_interests_end(line - 1)}; 2)"
      end

      def period_reimbursed_capitalized_interests_formula(line:)
        term = line - 1
        if term <= loan.deferred_and_capitalized + loan.deferred
          excel_float(float: 0.0)
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

      def delta_formula(line:)
        amount_added =
          if line == 2
            excel_float(float: 0.0)
          elsif line > 2
            "SOMME(#{column_range(column: amount_to_add, upto: line - 1)})"
          end
        "=ARRONDI(#{period_calculated_interests(line)} - #{period_interests(line)} - (#{capitalized_interests_end(line)} - #{capitalized_interests_start(line)}) - #{period_reimbursed_capitalized_interests(line)}; 14) - #{amount_added}"
      end

      def amount_to_add_formula(line:)
        "=TRONQUE(#{accrued_delta(line)}; 2)"
      end

      def row(term:)
        line = term + 1 # first term is on second line
        [
          term,
          loan.due_on + (term * loan.period_duration).months,
          remaining_capital_start_formula(line: line),
          remaining_capital_end_formula(line: line),
          period_calculated_interests_formula(line: line),
          delta_formula(line: line),
          period_accrued_delta_formula(line: line),
          amount_to_add_formula(line: line),
          period_interests_formula(line: line),
          period_capital_formula(line: line),
          total_paid_capital_end_of_period_formula(line: line),
          total_paid_interests_end_of_period_formula(line: line),
          period_total_formula(line: line),
          capitalized_interests_start_formula(line: line),
          capitalized_interests_end_formula(line: line),
          period_rate_formula,
          period_reimbursed_capitalized_interests_formula(line: line),
          period_calculated_capital_formula(line: line)
        ]
      end
    end
  end
end
