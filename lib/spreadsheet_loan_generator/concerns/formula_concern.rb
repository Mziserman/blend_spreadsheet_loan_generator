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
        "=#{remaining_capital_start(line)} * #{period_rate(line)}"
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
        "=-IPMT(#{period_rate(line)}; #{column_letter[:index]}#{line}; #{loan.duration}; #{excel_float(float: loan.amount)})"
      end

      def period_capital_formula(line:)
        "=-PPMT(#{period_rate(line)}; #{column_letter[:index]}#{line}; #{loan.duration}; #{excel_float(float: loan.amount)})"
      end

      def period_total_formula(line:)
        if loan.type == :standard
          "=-PMT(#{period_rate(line)}; #{loan.duration}; #{excel_float(float: loan.amount)})"
        else
          "=#{period_capital(line)} + #{period_interests(line)}"
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
          'capitalized_interests_start',
          'capitalized_interests_end',
          period_rate_formula
        ]
      end
    end
  end
end
