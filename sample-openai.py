import openai

# Set your API key here or via environment variable
openai.api_key = 'sk-proj-ebw0vkyVQ6CSgaMh9X0PT3BlbkFJqtwIYsIOllOnvjdGjwF1'

# Make a request to the OpenAI API
response = openai.Completion.create(
    engine="text-davinci-003",
    prompt="Say this is a test",
    max_tokens=5
)

# Print the response
print(response.choices[0].text.strip())
