module SpreadsheetLoanGenerator
  module SpreadsheetConcern
    extend ActiveSupport::Concern

    included do
      def columns
        %w[
          index
          due_on
          remaining_capital_start
          remaining_capital_end
          period_calculated_interests
          delta
          accrued_delta
          amount_to_add
          period_interests
          period_capital
          total_paid_capital_end_of_period
          total_paid_interests_end_of_period
          period_total
          capitalized_interests_start
          capitalized_interests_end
          period_rate
        ]
      end

      def column_letter
        {
          index: 'A',
          due_on: 'B',
          remaining_capital_start: 'C',
          remaining_capital_end: 'D',
          period_calculated_interests: 'E',
          delta: 'F',
          accrued_delta: 'G',
          amount_to_add: 'H',
          period_interests: 'I',
          period_capital: 'J',
          total_paid_capital_end_of_period: 'K',
          total_paid_interests_end_of_period: 'L',
          period_total: 'M',
          capitalized_interests_start: 'N',
          capitalized_interests_end: 'O',
          period_rate: 'P'
        }
      end

      def column_range(column: 'A', upto: , exclude_head: true)
        start_line = exclude_head ? 2 : 1

        "#{column}#{start_line}:#{column}#{upto}"
      end

      def index_to_line(index:)
        index + 1 # first term is on line 2
      end

      def excel_float(float:)
        float.to_s.gsub('.', ',')
      end

      def remaining_capital_start_formula(index:, amount:)
        base = excel_float(float: amount)

        return base if index == 2

        "=#{base} - #{column_letter[:total_paid_capital_end_of_period]}#{index - 1}"
      end

      def remaining_capital_end_formula(index:, amount:)
        base = excel_float(float: amount)

        "=#{base} - #{column_letter[:total_paid_capital_end_of_period]}#{index}"
      end

      def period_calculated_interests_formula(index:)
        "=#{column_letter[:remaining_capital_start]}#{index} * #{column_letter[:period_rate]}#{index}"
      end

      def period_accrued_delta_formula(index:)
        return excel_float(float: 0.0) if index == 2

        "=#{column_letter[:accrued_delta]}#{index - 1} + #{column_letter[:delta]}#{index}"
      end

      def total_paid_capital_end_of_period_formula(index:)
        return "=#{column_letter[:period_capital]}2" if index == 2

        "=SOMME(#{column_range(column: column_letter[:period_capital], upto: index)})"
      end

      def total_paid_interests_end_of_period_formula(index:)
        return "=#{column_letter[:period_interests]}2" if index == 2

        "=SOMME(#{column_range(column: column_letter[:period_interests], upto: index)})"
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
        "=-IPMT(#{column_letter[:period_rate]}#{index}; #{column_letter[:index]}#{index}; #{duration}; #{excel_float(float: amount)})"
      end

      def period_capital_formula(index:, duration:, amount:)
        "=-PPMT(#{column_letter[:period_rate]}#{index}; #{column_letter[:index]}#{index}; #{duration}; #{excel_float(float: amount)})"
      end

      def period_total_formula(index:, duration:, amount:, type: :standard)
        if type == :standard
          "=-PMT(#{column_letter[:period_rate]}#{index}; #{duration}; #{excel_float(float: amount)})"
        else
          "=#{column_letter[:period_capital]}#{index} + #{column_letter[:period_interests]}#{index}"
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
