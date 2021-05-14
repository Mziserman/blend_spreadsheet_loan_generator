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
        "=(#{remaining_capital_start(line)} + #{capitalized_interests_start(line)}) * #{period_rate(line)}"
      end

      def period_accrued_delta_formula(line:)
        return excel_float(float: 0.0) if line == 2

        "=#{accrued_delta(line - 1)} + #{delta(line)}"
      end

      def total_paid_capital_end_of_period_formula(line:)
        return "=#{period_capital(line)}" if line == 2

        "=SOMME(#{column_range(column: period_capital, upto: line)})"
      end

      def total_paid_interests_end_of_period_formula(line:)
        return "=#{period_interests(line)}" if line == 2

        "=SOMME(#{column_range(column: period_interests, upto: line)})"
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
          "=#{period_calculated_interests(line)}"
        else
          "=-IPMT(#{standard_params(line: line)})"
        end
      end

      def period_capital_formula(line:)
        "=#{period_calculated_capital(line)} - #{period_reimbursed_capitalized_interests(line)}"
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
        "=#{period_capital(line)} + #{period_interests(line)}"
      end

      def capitalized_interests_start_formula(line:)
        return excel_float(float: 0.0) if line == 2

        "=#{capitalized_interests_end(line - 1)}"
      end

      def period_reimbursed_capitalized_interests_formula(line:)
        term = line - 1
        if term <= loan.deferred_and_capitalized + loan.deferred
          excel_float(float: 0.0)
        else
          "=MIN(#{period_calculated_capital(line)}; #{capitalized_interests_start(line)})"
        end
      end

      def capitalized_interests_end_formula(line:)
        term = line - 1
        if term <= loan.deferred_and_capitalized
          "=#{capitalized_interests_start(line)} + #{period_calculated_interests(line)}"
        else
          "=#{capitalized_interests_start(line)} - #{period_reimbursed_capitalized_interests(line)}"
        end
      end

      def row(term:)
        line = term + 1 # first term is on second line
        [
          term,
          loan.due_on + (term * loan.period_duration).months,
          remaining_capital_start_formula(line: line),
          remaining_capital_end_formula(line: line),
          period_calculated_interests_formula(line: line),
          '',
          period_accrued_delta_formula(line: line),
          'amount_to_add',
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
