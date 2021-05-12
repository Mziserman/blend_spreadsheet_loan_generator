module SpreadsheetLoanGenerator
  module FormulaConcern
    extend ActiveSupport::Concern

    included do
      # use method missing defined in spreadsheet concerns a lot
      
      def remaining_capital_start_formula(index:, amount:)
        base = excel_float(float: amount)

        return base if index == 2

        "=#{base} - #{total_paid_capital_end_of_period(index - 1)}"
      end

      def remaining_capital_end_formula(index:, amount:)
        base = excel_float(float: amount)

        "=#{base} - #{total_paid_capital_end_of_period(index)}"
      end

      def period_calculated_interests_formula(index:)
        "=#{remaining_capital_start(index)} * #{period_rate(index)}"
      end

      def period_accrued_delta_formula(index:)
        return excel_float(float: 0.0) if index == 2

        "=#{accrued_delta(index - 1)} + #{delta(index)}"
      end

      def total_paid_capital_end_of_period_formula(index:)
        return "=#{period_capital(index)}" if index == 2

        "=SOMME(#{column_range(column: period_capital, upto: index)})"
      end

      def total_paid_interests_end_of_period_formula(index:)
        return "=#{period_interests(index)}" if index == 2

        "=SOMME(#{column_range(column: period_interests, upto: index)})"
      end

      def period_rate_formula(rate:, period_duration:, type: :simple)
        case type
        when :normal
          "=TAUX.NOMINAL(#{excel_float(float: rate)};#{excel_float(float: 12.0 / period_duration)})"
        when :simple
          "=#{excel_float(float: rate)} / 12,0"
        end
      end

      def period_interests_formula(index:, duration:, amount:)
        "=-IPMT(#{period_rate(index)}; #{column_letter[:index]}#{index}; #{duration}; #{excel_float(float: amount)})"
      end

      def period_capital_formula(index:, duration:, amount:)
        "=-PPMT(#{period_rate(index)}; #{column_letter[:index]}#{index}; #{duration}; #{excel_float(float: amount)})"
      end

      def period_total_formula(index:, duration:, amount:, type: :standard)
        if type == :standard
          "=-PMT(#{period_rate(index)}; #{duration}; #{excel_float(float: amount)})"
        else
          "=#{period_capital(index)} + #{period_interests(index)}"
        end
      end

      def line(index:, loan:)
        term = index + 1

        [
          term,
          loan.due_on + (term * loan.period_duration).months,
          remaining_capital_start_formula(index: index_to_line(index: term), amount: loan.amount),
          remaining_capital_end_formula(index: index_to_line(index: term), amount: loan.amount),
          period_calculated_interests_formula(index: index_to_line(index: term)),
          '',
          period_accrued_delta_formula(index: index_to_line(index: term)),
          'amount_to_add',
          period_interests_formula(index: index_to_line(index: term), duration: loan.duration, amount: loan.amount),
          period_capital_formula(index: index_to_line(index: term), duration: loan.duration, amount: loan.amount),
          total_paid_capital_end_of_period_formula(index: index_to_line(index: term)),
          total_paid_interests_end_of_period_formula(index: index_to_line(index: term)),
          period_total_formula(index: index_to_line(index: term), duration: loan.duration, amount: loan.amount),
          'capitalized_interests_start',
          'capitalized_interests_end',
          period_rate_formula(rate: loan.rate, period_duration: loan.period_duration)
        ]
      end
    end
  end
end
