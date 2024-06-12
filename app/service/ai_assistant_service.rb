class AiAssistantService
  require 'json'

  def self.call(informations, question)
    messages = [
      {
        role: 'system',
        content: JSON.pretty_generate(informations)
      },
      {
        role: 'system',
        content: 'Use the provided JSON to answer the following question.'
      }
    ]

    messages << { role: 'user', content: question }

    response = client.chat(
      parameters: {
          model: "gpt-3.5-turbo-16k",
          messages: messages,
          temperature: 0.4,
      })

    puts response.dig("choices", 0, "message", "content")
  end

  def self.client
    @client ||= OpenAI::Client.new(
      access_token: Rails.application.credentials.open_ai_api_key,
      log_errors: true
    )
  end
end