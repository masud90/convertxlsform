# outline of the all_types form in English

    Code
      print_outline(xlsform_outline(p, lang = "en"))
    Output
         row serial depth block       kind            name            label                                    required n_choices
      1   6  <NA>   0     note        note            intro_note      Household survey on water, food and a...    NA    NA       
      2   7  1      0     question    select_one      consent         Does the respondent agree to take par...  TRUE     2       
      3   8  2      0     question    text            follow_up       Enumerator code for this interview        TRUE    NA       
      4   9  3      0     section     begin_group     hh              Household                                   NA    NA       
      5  10  3.1    1     question    text            hh_code         Household code                            TRUE    NA       
      6  11  3.2    1     question    number          hh_size         Number of household members               TRUE    NA       
      7  12  3.3    1     question    number          land_ha         Area of land cultivated (hectares)       FALSE    NA       
      8  13  3.4    1     question    date            visit_date      Date of visit                            FALSE    NA       
      9  14  3.5    1     question    time            visit_time      Time of visit                            FALSE    NA       
      10 15  3.6    1     question    datetime        last_rain       Date and time of the last rainfall       FALSE    NA       
      11 16  3.7    1     question    select_one      water_source    Main source of drinking water            FALSE     4       
      12 17  3.8    1     question    select_multiple assets          Household assets                         FALSE     4       
      13 18  3.9    1     question    rank            priorities      Rank the household's priorities from ... FALSE     4       
      14 19  3.10   1     question    range           water_hours     Hours spent collecting water per day     FALSE    NA       
      15 20  3.11   1     question    range           food_share      Share of household income spent on fo... FALSE    NA       
      16 21  3.12   1     question    acknowledge     ack_understood  I confirm that the respondent underst... FALSE    NA       
      17 22  3.13   1     question    barcode         asset_tag       Scan the household asset tag             FALSE    NA       
      18 23  3.14   1     section     begin_group     water_food      Water and food                              NA    NA       
      19 24  3.14.1 2     question    select_multiple food_sources    Main sources of food in the past week    FALSE     4       
      20 25  3.14.2 2     question    text            food_note       Reason for not buying food at the market FALSE    NA       
      21 26  3.14.3 2     question    number          radio_hours     Hours per week the radio is used         FALSE    NA       
      22 29  3.15   2     question    number          asset_value     Estimated value of household assets (... FALSE    NA       
      23 30  3.16   2     question    select_one      owns_livestock  Does the household own livestock?        FALSE     2       
      24 32  3.17   1     question    number          livestock_count Number of livestock animals              FALSE    NA       
      25 34  4      0     section     begin_repeat    members         Household member                            NA    NA       
      26 35  4.1    1     question    text            member_name     First name of member                      TRUE    NA       
      27 36  4.2    1     question    number          member_age      Age of member (years)                    FALSE    NA       
      28 37  4.3    1     question    select_one      member_sex      Sex of member                            FALSE     2       
      29 38  4      0     section_end end_repeat      members         <NA>                                        NA    NA       
      30 39  5      0     question    geo             location        Location of the household (GPS)          FALSE    NA       
      31 40  6      0     question    geo             trace_route     Walking route to the water point         FALSE    NA       
      32 41  7      0     question    geo             plot_shape      Boundary of the cultivated plot          FALSE    NA       
      33 42  8      0     question    media           photo_dwelling  Photo of the dwelling                    FALSE    NA       
      34 43  9      0     question    media           audio_comment   Audio recording of the respondent's c... FALSE    NA       
      35 44  10     0     question    media           video_water     Video of the water point                 FALSE    NA       
      36 45  11     0     question    media           consent_form    Scanned consent form (PDF)               FALSE    NA       
      37 46  12     0     question    select_one      district        District                                 FALSE     3       
      38 47  13     0     question    select_one      village         Village (filtered by district)           FALSE     6       
      39 48  14     0     question    select_one      village_file    Village (from file)                      FALSE     5       
      40 49  15     0     question    select_multiple missing_items   Items from an external list              FALSE     3       
      41 50  16     0     question    select_one      main_crop       Main crop                                FALSE    30       
      42 52  17     0     question    text            comments        Additional comments                      FALSE    NA       
      43 53  <NA>   0     note        note            closing         Thank you. The survey is complete.          NA    NA       
         row serial relevant_text                      constraint_text                  
      6  11  3.2    <NA>                               Enter a whole number from 0 to 30
      20 25  3.14.2 Q3.14.1 does not include "Market"  <NA>                             
      21 26  3.14.3 Q3.2 > 2 and Q3.8 includes "Radio" <NA>                             
      22 29  3.15   number selected in Q3.8 > 1        <NA>                             
      24 32  3.17   Q3.16 is "Yes"                     <NA>                             
      27 36  4.2    <NA>                               Enter an age from 0 to 120       

# outline of the all_types form in French

    Code
      print_outline(xlsform_outline(p, lang = "fr"))
    Output
         row serial depth block       kind            name            label                                    required n_choices
      1   6  <NA>   0     note        note            intro_note      Enquête sur l'eau, l'alimentation et ...    NA    NA       
      2   7  1      0     question    select_one      consent         Le répondant accepte-t-il de particip...  TRUE     2       
      3   8  2      0     question    text            follow_up       Code de l'enquêteur pour cet entretien    TRUE    NA       
      4   9  3      0     section     begin_group     hh              Ménage                                      NA    NA       
      5  10  3.1    1     question    text            hh_code         Code du ménage                            TRUE    NA       
      6  11  3.2    1     question    number          hh_size         Nombre de membres du ménage               TRUE    NA       
      7  12  3.3    1     question    number          land_ha         Superficie cultivée (hectares)           FALSE    NA       
      8  13  3.4    1     question    date            visit_date      Date de la visite                        FALSE    NA       
      9  14  3.5    1     question    time            visit_time      Heure de la visite                       FALSE    NA       
      10 15  3.6    1     question    datetime        last_rain       Date et heure de la dernière pluie       FALSE    NA       
      11 16  3.7    1     question    select_one      water_source    Principale source d'eau potable          FALSE     4       
      12 17  3.8    1     question    select_multiple assets          Biens du ménage                          FALSE     4       
      13 18  3.9    1     question    rank            priorities      Classer les priorités du ménage de la... FALSE     4       
      14 19  3.10   1     question    range           water_hours     Heures passées à chercher de l'eau pa... FALSE    NA       
      15 20  3.11   1     question    range           food_share      Part du revenu du ménage dépensée en ... FALSE    NA       
      16 21  3.12   1     question    acknowledge     ack_understood  Je confirme que le répondant a compri... FALSE    NA       
      17 22  3.13   1     question    barcode         asset_tag       Scanner l'étiquette du bien du ménage    FALSE    NA       
      18 23  3.14   1     section     begin_group     water_food      Eau et alimentation                         NA    NA       
      19 24  3.14.1 2     question    select_multiple food_sources    Principales sources d'alimentation de... FALSE     4       
      20 25  3.14.2 2     question    text            food_note       Raison de ne pas acheter de la nourri... FALSE    NA       
      21 26  3.14.3 2     question    number          radio_hours     Heures par semaine d'utilisation de l... FALSE    NA       
      22 29  3.15   2     question    number          asset_value     Valeur estimée des biens du ménage (USD) FALSE    NA       
      23 30  3.16   2     question    select_one      owns_livestock  Le ménage possède-t-il du bétail ?       FALSE     2       
      24 32  3.17   1     question    number          livestock_count Nombre d'animaux d'élevage               FALSE    NA       
      25 34  4      0     section     begin_repeat    members         Membre du ménage                            NA    NA       
      26 35  4.1    1     question    text            member_name     Prénom du membre                          TRUE    NA       
      27 36  4.2    1     question    number          member_age      Âge du membre (années)                   FALSE    NA       
      28 37  4.3    1     question    select_one      member_sex      Sexe du membre                           FALSE     2       
      29 38  4      0     section_end end_repeat      members         <NA>                                        NA    NA       
      30 39  5      0     question    geo             location        Localisation du ménage (GPS)             FALSE    NA       
      31 40  6      0     question    geo             trace_route     Itinéraire à pied jusqu'au point d'eau   FALSE    NA       
      32 41  7      0     question    geo             plot_shape      Limite de la parcelle cultivée           FALSE    NA       
      33 42  8      0     question    media           photo_dwelling  Photo du logement                        FALSE    NA       
      34 43  9      0     question    media           audio_comment   Enregistrement audio des commentaires... FALSE    NA       
      35 44  10     0     question    media           video_water     Vidéo du point d'eau                     FALSE    NA       
      36 45  11     0     question    media           consent_form    Formulaire de consentement scanné (PDF)  FALSE    NA       
      37 46  12     0     question    select_one      district        District                                 FALSE     3       
      38 47  13     0     question    select_one      village         Village (filtré par district)            FALSE     6       
      39 48  14     0     question    select_one      village_file    Village (depuis un fichier)              FALSE     5       
      40 49  15     0     question    select_multiple missing_items   Éléments d'une liste externe             FALSE     3       
      41 50  16     0     question    select_one      main_crop       Culture principale                       FALSE    30       
      42 52  17     0     question    text            comments        Commentaires supplémentaires             FALSE    NA       
      43 53  <NA>   0     note        note            closing         Merci. L'enquête est terminée.              NA    NA       
         row serial relevant_text                              constraint_text                  
      6  11  3.2    <NA>                                       Enter a whole number from 0 to 30
      20 25  3.14.2 Q3.14.1 ne comprend pas "Marché"           <NA>                             
      21 26  3.14.3 Q3.2 > 2 et Q3.8 comprend "Radio"          <NA>                             
      22 29  3.15   nombre de choix sélectionnés dans Q3.8 > 1 <NA>                             
      24 32  3.17   Q3.16 est "Oui"                            <NA>                             
      27 36  4.2    <NA>                                       Enter an age from 0 to 120       

# outline of the multilang form in English

    Code
      print_outline(xlsform_outline(p, lang = "en"))
    Output
         row serial depth block       kind            name                 label                                     required n_choices
      1   5  1      0     section     begin_group     groupA               A: Interview Metadata                       NA      NA      
      2   6  1.1    1     question    date            surveydate           What is the date today?                   TRUE      NA      
      3   7  1.2    1     question    select_one      enumname             Select enumerator name.                   TRUE     100      
      4  10  <NA>   1     note        note            note_consent         **Consent message**:\nThis message wil...   NA      NA      
      5  11  1.3    1     question    acknowledge     consent              Has the respondent consented?             TRUE      NA      
      6  13  2      0     section     begin_group     groupB               B: Household roster                         NA      NA      
      7  14  2.1    1     question    number          hhsize               How many individuals do you have in y...  TRUE      NA      
      8  15  2.2    1     section     begin_repeat    hhroster             B1. Household member details                NA      NA      
      9  17  2.2.1  2     question    text            mem_name             What is the name of person [hhroster_...  TRUE      NA      
      10 18  2.2.2  2     question    select_one      mem_sex              What is the sex of <span style="color...  TRUE       2      
      11 19  2.2.3  2     question    date            mem_birthyear        In which year was <span style="color:...  TRUE      NA      
      12 21  2.2.4  2     question    select_one      mem_relationship     What is the relationship of <span sty...  TRUE       7      
      13 22  2.2    1     section_end end_repeat      hhroster             <NA>                                        NA      NA      
      14 24  3      0     section     begin_group     groupC               C: Housing/ Shelter                         NA      NA      
      15 25  3.1    1     question    geo             location_coordinates What are the coordinates of this dwel...  TRUE      NA      
      16 26  3.2    1     question    media           landmark_photo       Please take a photo of the house.         TRUE      NA      
      17 27  3.3    1     question    media           landmark_video       Please take a video of the street in ...  TRUE      NA      
      18 28  3.4    1     question    number          room_num             How many rooms are there in this house?   TRUE      NA      
      19 29  3.5    1     question    select_one      dwelling_ownership   Do you own this dwelling?                 TRUE       2      
      20 30  3.6    1     question    select_one      dwelling_rent        Do you pay rent to stay in this dwell...  TRUE       2      
      21 31  3.7    1     question    range           rent_range           How much do you pay as rent for this ...  TRUE      NA      
      22 33  4      0     section     begin_group     groupD               D: Assets                                   NA      NA      
      23 34  4.1    1     question    select_multiple asset_owned          Which of the following assets do you ...  TRUE      10      
      24 35  4.2    1     question    text            asset_owned_other    If you selected Other, specify in the...  TRUE      NA      
      25 36  4.3    1     section     begin_group     groupD1              D1: Mobile Phone                            NA      NA      
      26 37  4.3.1  2     question    select_one      phone_brand          What is the name of your phone manufa...  TRUE       7      
      27 38  4.3.2  2     question    text            phone_brand_other    If you selected Other, specify in the...  TRUE      NA      
      28 39  4.3.3  2     question    rank            phone_feature        Which features of your phone are the ...  TRUE       5      
      29 42  5      0     section     begin_group     groupE               E: Some more test questions                 NA      NA      
      30 43  5.1    1     question    number          decimal_entry        What is a number that is bigger than ...  TRUE      NA      
      31 44  5.2    1     question    geo             street_coordinates   Please trace the coordinates of this ...  TRUE      NA      
      32 45  5.3    1     question    geo             house_coordinates    Please draw a polygon shape of the ar...  TRUE      NA      
      33 46  5.4    1     question    barcode         barcode_entry        Scan the barcode of a book in front o...  TRUE      NA      
         row serial relevant_text                    constraint_text                                                                                                   
      2   6  1.1    <NA>                             Selected date must be the same as today.                                                                          
      7  14  2.1    <NA>                             Entered number must be higher than 0. Maximum allowed number is 30.                                               
      11 19  2.2.3  <NA>                             Year cannot be greater than 2025, or lower than 1900.                                                             
      20 30  3.6    Q3.5 is "Yes"                    <NA>                                                                                                              
      21 31  3.7    Q3.6 is "Yes"                    <NA>                                                                                                              
      23 34  4.1    <NA>                             If you select **None of these** or **Refused to respond** or **Do not know**, you can not select any other option.
      24 35  4.2    Q4.1 includes "Other, specify"   <NA>                                                                                                              
      25 36  4.3    Q4.1 includes "Mobile phone"     <NA>                                                                                                              
      27 38  4.3.2  Q4.3.1 includes "Other, specify" <NA>                                                                                                              
      30 43  5.1    <NA>                             the answer > 1 and the answer < 2                                                                                 

# outline of the multilang form in Bengali

    Code
      print_outline(xlsform_outline(p, lang = "bn"))
    Output
         row serial depth block       kind            name                 label                                    required n_choices
      1   5  1      0     section     begin_group     groupA               A: সাক্ষাৎকারের মেটাডেটা                    NA      NA      
      2   6  1.1    1     question    date            surveydate           আজ কত তারিখ?                             TRUE      NA      
      3   7  1.2    1     question    select_one      enumname             গণনাকারীর নাম নির্বাচন করুন।               TRUE     100      
      4  10  <NA>   1     note        note            note_consent         **সম্মতি বার্তা**:\nএই বার্তাটি [enumn...     NA      NA      
      5  11  1.3    1     question    acknowledge     consent              নির্বাচিত উত্তরদাতা কি সম্মতি দিয়েছেন?      TRUE      NA      
      6  13  2      0     section     begin_group     groupB               B: পরিবারের তালিকা                         NA      NA      
      7  14  2.1    1     question    number          hhsize               আপনার পরিবারে/ খানায় কতজন ব্যক্তি আছেন?    TRUE      NA      
      8  15  2.2    1     section     begin_repeat    hhroster             B1. পরিবারের সদস্যদের বিবরণ                 NA      NA      
      9  17  2.2.1  2     question    text            mem_name             সদস্য সংখ্যা  [hhroster_position] এর ...   TRUE      NA      
      10 18  2.2.2  2     question    select_one      mem_sex              <span style="color:red">[Q2.2.1]</spa... TRUE       2      
      11 19  2.2.3  2     question    date            mem_birthyear        <span style="color:red">[Q2.2.1]</spa... TRUE      NA      
      12 21  2.2.4  2     question    select_one      mem_relationship     <span style="color:red">[Q2.2.1]</spa... TRUE       7      
      13 22  2.2    1     section_end end_repeat      hhroster             <NA>                                       NA      NA      
      14 24  3      0     section     begin_group     groupC               C: আবাসন/আশ্রয়স্থল                           NA      NA      
      15 25  3.1    1     question    geo             location_coordinates এই বাসস্থানের স্থানাঙ্কগুলি কী কী?           TRUE      NA      
      16 26  3.2    1     question    media           landmark_photo       দয়া করে বাড়ির একটা ছবি তুলুন।              TRUE      NA      
      17 27  3.3    1     question    media           landmark_video       বাড়ির সামনের রাস্তার একটা ভিডিও করুন।      TRUE      NA      
      18 28  3.4    1     question    number          room_num             এই বাড়িতে কয়টি ঘর আছে?                   TRUE      NA      
      19 29  3.5    1     question    select_one      dwelling_ownership   আপনি কি এই বাড়ির মালিক?                  TRUE       2      
      20 30  3.6    1     question    select_one      dwelling_rent        আপনি এই বাড়িতে থাকার জন্য কি ভাড়া দেন ?   TRUE       2      
      21 31  3.7    1     question    range           rent_range           আপনি এই বাড়িতে থাকার জন্য কত টাকা ভাড়...  TRUE      NA      
      22 33  4      0     section     begin_group     groupD               D: সম্পদ                                    NA      NA      
      23 34  4.1    1     question    select_multiple asset_owned          আপনি নিচের কোন সম্পদের মালিক?             TRUE      10      
      24 35  4.2    1     question    text            asset_owned_other    যদি আপনি আগের প্রশ্নে "অন্যান্য, উল্ল...      TRUE      NA      
      25 36  4.3    1     section     begin_group     groupD1              D1: মোবাইল ফোন                             NA      NA      
      26 37  4.3.1  2     question    select_one      phone_brand          আপনার ফোন প্রস্তুতকারকের নাম কি?            TRUE       7      
      27 38  4.3.2  2     question    text            phone_brand_other    যদি আপনি আগের প্রশ্নে "অন্যান্য, উল্ল...      TRUE      NA      
      28 39  4.3.3  2     question    rank            phone_feature        আপনার ফোনের কোন বৈশিষ্ট্যগুলি আপনার ফ...    TRUE       5      
      29 42  5      0     section     begin_group     groupE               E: আরও কিছু প্রশ্ন                            NA      NA      
      30 43  5.1    1     question    number          decimal_entry        এমন একটি সংখ্যা কী যা ১ এর চেয়ে বড় ...    TRUE      NA      
      31 44  5.2    1     question    geo             street_coordinates   এই রাস্তার স্থানাঙ্কগুলি দাগান।              TRUE      NA      
      32 45  5.3    1     question    geo             house_coordinates    বাড়ির ক্ষেত্রফলের একটি বহুভুজ আকৃতি ...       TRUE      NA      
      33 46  5.4    1     question    barcode         barcode_entry        আপনার সামনে থাকা একটি বইয়ের বারকোড স...  TRUE      NA      
         row serial relevant_text                       constraint_text                                                                                                                     
      2   6  1.1    <NA>                                নির্বাচিত তারিখটি আজকের তারিখ হতে হবে।                                                                                               
      7  14  2.1    <NA>                                নির্বাচিত সংখ্যাটি অবশ্যই ০ এর বেশি হতে হবে। সর্বোচ্চ অনুমোদিত সংখ্যা হচ্ছে ৩০।                                                             
      11 19  2.2.3  <NA>                                বছরটি ২০২৫ সালের বেশি বা ১৯০০ সালের কম হতে পারবে না।                                                                                
      20 30  3.6    Q3.5 is "হ্যাঁ"                       <NA>                                                                                                                                
      21 31  3.7    Q3.6 is "হ্যাঁ"                       <NA>                                                                                                                                
      23 34  4.1    <NA>                                যদি আপনি **এর কোনটিই নয়** অথবা **উত্তর দিতে অস্বীকৃতি** অথবা **জানি না** নির্বাচন করেন, তাহলে আপনি অন্য কোন বিকল্প নির্বাচন করতে পারবেন না।
      24 35  4.2    Q4.1 includes "অন্যান্য, উল্লেখ করুন"   <NA>                                                                                                                                
      25 36  4.3    Q4.1 includes "মোবাইল ফোন"          <NA>                                                                                                                                
      27 38  4.3.2  Q4.3.1 includes "অন্যান্য, উল্লেখ করুন" <NA>                                                                                                                                
      30 43  5.1    <NA>                                the answer > 1 and the answer < 2                                                                                                   

