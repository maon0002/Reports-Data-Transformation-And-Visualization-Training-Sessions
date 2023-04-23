import pandas as pd


class Transliteration:

    @staticmethod
    def transliterate_bg_to_en(df: pd.DataFrame, ser: pd.Series, new_ser: pd.Series):

        """
        Transliterate bg chars into latin when needed
        """
        # create dictionary for transliteration from bg to latin letters
        bg_en_dict = {
            'А': 'A',
            'Б': 'B',
            'В': 'V',
            'Г': 'G',
            'Д': 'D',
            'Е': 'E',
            'Ж': 'Zh',
            'З': 'Z',
            'И': 'I',
            'Й': 'Y',
            'К': 'K',
            'Л': 'L',
            'М': 'M',
            'Н': 'N',
            'О': 'O',
            'П': 'P',
            'Р': 'R',
            'С': 'S',
            'Т': 'T',
            'У': 'U',
            'Ф': 'F',
            'Х': 'H',
            'Ц': 'Ts',
            'Ч': 'Ch',
            'Ш': 'Sh',
            'Щ': 'Sht',
            'Ъ': 'A',
            'Ь': 'Y',
            'Ю': 'Yu',
            'Я': 'Ya',
            'а': 'a',
            'б': 'b',
            'в': 'v',
            'г': 'g',
            'д': 'd',
            'е': 'e',
            'ж': 'zh',
            'з': 'z',
            'и': 'i',
            'й': 'y',
            'к': 'k',
            'л': 'l',
            'м': 'm',
            'н': 'n',
            'о': 'o',
            'п': 'p',
            'р': 'r',
            'с': 's',
            'т': 't',
            'у': 'u',
            'ф': 'f',
            'х': 'h',
            'ц': 'ts',
            'ч': 'ch',
            'ш': 'sh',
            'щ': 'sht',
            'ъ': 'a',
            'ь': 'y',
            'ю': 'yu',
            'я': 'ya',
            ' ': ' ',
            '-': '-'}

        # transliteration char by char
        transliterated_string = ""
        for i in range(len(df[ser])):
            string_value = df.loc[i, ser].strip()
            if string_value and not string_value.isascii():
                for char in string_value:
                    if char in bg_en_dict.keys():
                        string_value = string_value.replace(
                            char, bg_en_dict[char])
                        transliterated_string = string_value
                    else:
                        transliterated_string = string_value

                df.loc[i, new_ser] = transliterated_string.upper()
                transliterated_string = ""
            else:
                df.loc[i, new_ser] = df.loc[i, ser].upper()
        return df[new_ser]
