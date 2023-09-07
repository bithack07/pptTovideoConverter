import openai
import configparser
file_path = 'chatGPTprops.properties'



#To read the chat gpt properties file
def read_properties_file(file_path):
    config = configparser.ConfigParser()
    config.read(file_path)
    
    properties = {}
    
    for section in config.sections():
        for key, value in config.items(section):
            properties[key] = value
    
    return properties

properties = read_properties_file(file_path)
openai.api_key= properties.get("api_key")
openai.api_base = properties.get("endpoint")
openai.api_type = properties.get("api_type")
openai.api_version = properties.get("api_version")

def get_response(content,prompt):

    
    try:
        response=openai.ChatCompletion.create(
            engine="gpt-35-turbo",
            n=1,
            stop=None,
            temperature=0,
            messages=[
                {"role": "system", "content": content},
                {"role": "user", "content": prompt}
            ]
        )
        responseData=response.choices[0].message.content.strip()
        #print(responseData)
    except Exception:
        responseData = None
    return responseData 


def refine_slide_text(slideText):
    #slide= slideText.strip()
    slide_numbers = slideText.split('\n')[0].strip()
    slide_number= slide_numbers.replace(":", "")
    print(f"GPT is refining {slide_number} ")
    content= "You are a content formatter. Carefully analyze the content provided and refine the content."
    
    prompt= slideText
    refined_slideData = get_response(content,prompt)
    #print("Done")
    return refined_slideData


def generate_voice_over_scripts(refined_texts):
    slide_numbers = refined_texts.split(":")[0].strip()
    slide_number= slide_numbers.replace(":", "")
    print(f"GPT is generating voiceovers for  {slide_number} ")
    content= "Carefully analyze the content provided and make professional voice-over scripts and keep the responses formatted as '<content>'."
    prompt= refined_texts
    refined_slideData = get_response(content,prompt)

    #print("Done")
    return refined_slideData




