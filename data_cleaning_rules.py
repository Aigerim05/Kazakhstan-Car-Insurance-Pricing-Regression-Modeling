import pandas as pd
from dateutil.relativedelta import relativedelta


class InsuranceDataCleaner:
    def __init__(self, path):
        self.df = pd.read_excel(path)

    def drop_features(self):
        """
        Drops unnecessary columns from the dataset.
        """
        self.df.drop(columns=['Unique number', 'Citizenship', 'Gender', 'Loss_amount', 'Accident_region'], inplace=True)


    def adjust_invalid_driving_experience(self):
        """ Removes rows with invalid driving experience and adjusts cases 
            where a person started driving before age 18.
        """

        # Remove rows where driving experience is more than age
        self.df['Age_gap'] = self.df['Age'] - self.df['Driving_experience']
        invalid_experience = self.df[self.df['Driving_experience'] > self.df['Age']]
        self.df.drop(index=invalid_experience.index, inplace=True)

        # Adjust experience if the person started driving before age 18
        started_under_18 = self.df[self.df['Age_gap'] < 18]
        adjustment = 18 - started_under_18['Age_gap']
        self.df.loc[started_under_18.index, 'Driving_experience'] -= adjustment

        # Drop helper column
        self.df.drop(columns=['Age_gap'], inplace=True)

        # remove negative values (if still present)
        self.df = self.df[self.df['Driving_experience'] >= 0]

    

    def clean_vehicle_and_insurance_info(self):
        """
        Cleans vehicle type categories, calculates car age from manufacturing year, 
        and extracts insurance period in months.
        """

        # Clean column names
        self.df.columns = self.df.columns.str.strip()

        # Filter and map vehicle types
        self.df = self.df[~self.df['Vehicle_type'].isin([
            'Прицеп к грузовой а/м',
            'Прицеп к легковой а/м'
        ])]

        vehicle_mapping = {
            'Легковые автомобили': 'Легковые автомобили',
            'Мотоциклы и мотороллеры': 'Мотоциклы',
            'Грузовые автомобили': 'Грузовые',
            'Автобусы до 16 п/м вкл.': 'Автобусы',
            'Автобусы, свыше 16 п/м': 'Автобусы'
        }

        self.df['Vehicle_type'] = self.df['Vehicle_type'].map(vehicle_mapping)

        # Calculate car age
        self.df['Year_of_manufacture'] = pd.to_datetime(self.df['Year_of_manufacture'], errors='coerce')
        self.df['Car_age'] = 2025 - self.df['Year_of_manufacture'].dt.year
        self.df.drop(columns=['Year_of_manufacture'], inplace=True)


    def split_insurance_dates(self):
        """
        Splits the 'Insurance_period' column into 'start_date' and 'end_date',
        converting them to datetime format.
        """
        self.df[['start_date', 'end_date']] = self.df['Insurance_period'].str.split('-', expand=True)
        self.df['start_date'] = pd.to_datetime(self.df['start_date'], format='%d.%m.%Y', errors='coerce')
        self.df['end_date'] = pd.to_datetime(self.df['end_date'], format='%d.%m.%Y', errors='coerce')

   
    def calculate_insurance_duration(self):
        """
        Calculates the insurance duration in months using 'start_date' and 'end_date',
        and stores the result in 'Insurance_months'.
        Drops the intermediate date columns afterward.
        """

        def calculate_months(row):
            diff = relativedelta(row['end_date'], row['start_date'])
            return diff.years * 12 + (diff.months + 1)

        self.df['Insurance_months'] = self.df.apply(calculate_months, axis=1)
        self.df.drop(columns=['Insurance_period', 'start_date', 'end_date'], inplace=True)


    def fill_missing_privileges(self):
        """
        Fills missing values in the 'Privileges' column with 'Не инвалид'.
        """
        if 'Privileges' in self.df.columns:
            self.df['Privileges'] = self.df['Privileges'].fillna('Не инвалид')


    def standardize_colors(self):
        """
        Standardizes the 'Color' column by grouping similar shades under common categories.
        Unrecognized or rare colors are labeled as 'Прочие'.
        """
        if 'Color' not in self.df.columns:
            return

        color_mapping = {
            'белый': 'Белый', 'снежная королева': 'Белый', 'жемчужно-белый': 'Белый', 'бело-серый': 'Белый',
            'черный': 'Черный', 'черный металлик': 'Черный', 'черный с фиолетовым отливом': 'Черный',
            'серый': 'Серый', 'темно-серый': 'Серый', 'графит': 'Серый', 'мокрый асфальт': 'Серый',
            'синий': 'Синий', 'темно-синий': 'Синий', 'ярко-синий': 'Синий', 'сине-зеленый': 'Синий',
            'зеленый': 'Зеленый', 'лайм': 'Зеленый', 'изумрудный': 'Зеленый', 'оливковый': 'Зеленый',
            'красный': 'Красный', 'бордовый': 'Красный', 'вишневый': 'Красный', 'коралл': 'Красный',
            'желтый': 'Желтый', 'лимонный': 'Желтый', 'светло-желтый': 'Желтый',
            'коричневый': 'Коричневый', 'мокко': 'Коричневый', 'шоколадный': 'Коричневый',
            'фиолетовый': 'Фиолетовый', 'сиреневый': 'Фиолетовый', 'лиловый': 'Фиолетовый',
        }

        self.df['Color'] = self.df['Color'].astype(str).apply(
            lambda col: color_mapping.get(col.strip().lower(), 'Прочие')
        )

    def standardize_brands_and_models(self):
        """
        Standardizes car brands and models by:
        - Replacing 'Лада' with 'Lada'
        - Grouping less frequent brands into 'Other' (keeps top 37)
        - Replacing '.' in 'Model' with 'Unknown'
        - Grouping less frequent models into 'Other' (keeps top 50)
        """
        if 'Brand' in self.df.columns:
            self.df['Brand'] = self.df['Brand'].replace({'Лада': 'Lada'})
            top_brands = self.df['Brand'].value_counts().nlargest(37).index
            self.df['Brand'] = self.df['Brand'].apply(lambda b: b if b in top_brands else 'Other')

        if 'Model' in self.df.columns:
            self.df['Model'] = self.df['Model'].replace('.', 'Unknown')
            top_models = self.df['Model'].value_counts().nlargest(50).index
            self.df['Model'] = self.df['Model'].apply(lambda m: m if m in top_models else 'Other')


    def map_city_to_region(self):
        """
        Cleans the 'City' column by extracting the main city name (before any comma),
        maps it to a region using a predefined dictionary, and stores the result in a new 'Region' column.
        Drops intermediate columns and fills missing values as 'Unknown'.
        """
        if 'City' not in self.df.columns:
            return

        # Extract main city name
        self.df['City_cleaned'] = self.df['City'].astype(str).apply(lambda x: x.split(',')[0].strip())

        # Mapping dictionary
        region_mapping = {
            'Алматы': 'Алматинская область',
            'Нур-Султан': 'Астана',
            'Актобе': 'Актюбинская область',
            'Петропавловск': 'Северо-Казахстанская область',
            'Кокшетау': 'Акмолинская область',
            'Костанай': 'Костанайская область',
            'Павлодар': 'Павлодарская область',
            'Караганда': 'Карагандинская область',
            'Семей': 'Восточно-Казахстанская область',
            'Актау': 'Мангистауская область',
            'Атырау': 'Атырауская область',
            'Уральск': 'Западно-Казахстанская область',
            'Талдыкорган': 'Алматинская область',
            'Шымкент': 'Туркестанская область',
            'Кызылорда': 'Кызылординская область',
            'Тараз': 'Жамбылская область',
            'Усть-Каменогорск': 'Восточно-Казахстанская область',
            'Рудный': 'Костанайская область',
            'Темиртау': 'Карагандинская область',
            'Есик': 'Алматинская область',
            'Атбасар': 'Акмолинская область',
            'Жаксы': 'Акмолинская область',
            'Есиль': 'Северо-Казахстанская область',
            'Красный Яр': 'Актюбинская область',
            'Мариновка': 'Костанайская область',
            'Запорожье': 'Костанайская область',
            'Новоалександровка': 'Костанайская область',
            'Балкашино': 'Акмолинская область',
            'Лозовое': 'Костанайская область',
            'Талгар': 'Алматинская область',
            'Есенгельды': 'Алматинская область',
            'Новокиенка': 'Костанайская область',
            'Борисовка': 'Северо-Казахстанская область',
            'Аршалы': 'Акмолинская область',
            'Максимовка': 'Акмолинская область',
            'Боралдай': 'Алматинская область',
            'Покровка': 'Актюбинская область',
            'Октябрьское': 'Актюбинская область',
            'Садовое': 'Северо-Казахстанская область',
            'Тимашевка': 'Костанайская область'
        }

        # Map city to region
        self.df['Region'] = self.df['City_cleaned'].map(region_mapping)

        # Drop helper columns
        self.df.drop(columns=["City_cleaned", "City"], inplace=True)

        # Fill unmapped regions
        self.df['Region'] = self.df['Region'].fillna('Unknown')


    def clean(self):
        """
        Runs the full cleaning pipeline on the dataset.
        Applies all preprocessing steps in a defined order.
        Returns the cleaned DataFrame.
        """
        self.drop_features()
        self.adjust_invalid_driving_experience()
        self.clean_vehicle_and_insurance_info()
        self.split_insurance_dates()
        self.calculate_insurance_duration()
        self.fill_missing_privileges()
        self.standardize_colors()
        self.standardize_brands_and_models()
        self.map_city_to_region()
        return self.df
