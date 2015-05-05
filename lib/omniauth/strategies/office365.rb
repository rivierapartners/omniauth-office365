require 'omniauth-oauth2'

module OmniAuth
  module Strategies
    class Office365 < OmniAuth::Strategies::OAuth2
      option :client_options, {
                                site: 'https://outlook.office365.com/',
                                token_url: 'https://login.windows.net/rivierapartners.com/oauth2/token',
                                authorize_url: 'https://login.windows.net/rivierapartners.com/oauth2/authorize'
                            }

      def request_phase
        super
      end

      def authorize_params
        params = super
        params[:resource] = 'https://outlook.office365.com/'
        params
      end

      uid { raw_info["MailboxGuid"] }

      info do
        {
            'email' => raw_info["Id"],
            'name' => raw_info["DisplayName"],
            'nickname' => raw_info["Alias"],
            'first_name' => raw_info["first_name"],
            'last_name' => raw_info["last_name"]
        }
      end

      def raw_info
        @raw_info ||= access_token.get('/ews/odata/me').parsed
        names = @raw_info["DisplayName"].split(' ')
        @raw_info["first_name"] = names.first
        @raw_info["last_name"] = names[1..-1].join(' ')
        @raw_info
      end


    end
  end
end

OmniAuth.config.add_camelization 'office365', 'Office365'
