module SpreadsheetLoanGenerator
  module FormulaConcern
    extend ActiveSupport::Concern

    included do
      # use method missing defined in spreadsheet concerns a lot

      def remaining_capital_start_formula(line:, amount:)
        base = excel_float(float: amount)

        return base if line == 2

        "=#{base} - #{total_paid_capital_end_of_period(line - 1)}"
      end

      def remaining_capital_end_formula(line:, amount:)
        base = excel_float(float: amount)

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

      def period_rate_formula(rate:, period_duration:, type: :simple)
        case type
        when :normal
          "=TAUX.NOMINAL(#{excel_float(float: rate)};#{excel_float(float: 12.0 / period_duration)})"
        when :simple
          "=#{excel_float(float: rate)} / 12,0"
        end
      end

      def period_interests_formula(line:, duration:, amount:)
        "=-IPMT(#{period_rate(line)}; #{column_letter[:index]}#{line}; #{duration}; #{excel_float(float: amount)})"
      end

      def period_capital_formula(line:, duration:, amount:)
        "=-PPMT(#{period_rate(line)}; #{column_letter[:index]}#{line}; #{duration}; #{excel_float(float: amount)})"
      end

      def period_total_formula(line:, duration:, amount:, type: :standard)
        if type == :standard
          "=-PMT(#{period_rate(line)}; #{duration}; #{excel_float(float: amount)})"
        else
          "=#{period_capital(line)} + #{period_interests(line)}"
        end
      end

      def row(term:, loan:)
        line = term + 1 # first term is on second line
        [
          term,
          loan.due_on + (term * loan.period_duration).months,
          remaining_capital_start_formula(line: line, amount: loan.amount),
          remaining_capital_end_formula(line: line, amount: loan.amount),
          period_calculated_interests_formula(line: line),
          '',
          period_accrued_delta_formula(line: line),
          'amount_to_add',
          period_interests_formula(line: line, duration: loan.duration, amount: loan.amount),
          period_capital_formula(line: line, duration: loan.duration, amount: loan.amount),
          total_paid_capital_end_of_period_formula(line: line),
          total_paid_interests_end_of_period_formula(line: line),
          period_total_formula(line: line, duration: loan.duration, amount: loan.amount),
          'capitalized_interests_start',
          'capitalized_interests_end',
          period_rate_formula(rate: loan.rate, period_duration: loan.period_duration)
        ]
      end
    end
  end
end
