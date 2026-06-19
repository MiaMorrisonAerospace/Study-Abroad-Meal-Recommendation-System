% ------------------------------------------------------------------------
%   Name: Mia Morrison
%   Section: 05
%   Finish Date: 4/28/2024
%
%  [ This script is a recipe generator. The user will
%   input what meal, meal type(protein), and how much time they have.
%   Matlab will use the forloop to search the attached excel sheet with this
%   informtion and the website with that recipe will populate on their
%   screen. In addition, there will be a bar chart behind it, which shows
%   the Macros: the amount of calories, protein, carbs, and fat in grams.]
%
%
%   Citation: [
%
%                        Resources  |  1-3
%        https://www.mathworks.com/help/matlab/ref/web.html
%                                N/A|  117
%
%
%                        Resources  |  12-14
%        https://www.mathworks.com/matlabcentral/answers/593683-how-to-
%                    add-newline-to-barplot-x-label
%                                N/A|  124
%
%   CORE TECHNIQUES:
%
%   <SM:ROP:MORRISON>        62
%   <SM:BOP:MORRISON>        62
%   <SM:IF:MORRISON>         70
%   <SM:FOR:MORRISON>        7 inside the pdf
%   <SM:WHILE:MORRISON>      64
%   <SM:RANDOM:MORRISON>     77,84,99
%   <SM:PDF:MORRISON>        111
%   <SM:PDF_PARAM:MORRISON>  57,70,71
%   <SM:PDF_RETURN:MORRISON> 120
%   <SM:READ:MORRISON>       57
%   
%   SPECIFIC TECHNIQUES:
%
%   <SM:STRING:MORRISON>     54,169
%   <SM:REF:MORRISON>        147
%   <SM:SLICE:MORRISON>      9 inside the pdf
%   <SM:SEARCH:MORRISON>     8 inside the pdf
%   <SM:VIEW:MORRISON>       166
%   <SM:PLOT:MORRISON>       122,134,147
% ]
% -------------------------------------------------------------------------

% INTRODUCTION

clear; clc; close all
Run_again = 'yes';
while strcmpi(Run_again,'yes')==1   %<SM:STRING:MORRISON>  %To ensure that the user can run the code again
    clear; clc; close all

    Recipes = readcell('Recipes.xlsx'); %<SM:PDF_PARAM:MORRISON> %<SM:READ:MORRISON>


    fprintf('\nWelcome to ''Chef''s Kiss'', your healthy, high protein, low calorie recipe generator.\n')
    Meal = input('\nWhat meal would you like to eat? (1 = Breakfast, 2 = Lunch, 3 = Dinner, 4 = Dessert) ');
    while isempty(Meal)==1 || Meal < 1 || Meal > 4 || mod(Meal,1) ~= 0 %<SM:WHILE:MORRISON>  %<SM:ROP:MORRISON>  %<SM:BOP:MORRISON>
        Meal = input('ERROR. What meal would you like to eat? (1 - Breakfast, 2 - Lunch, 3 - Dinner, 4 - Dessert) \n');
    end



    % INPUTS

    if Meal == 1 %<SM:IF:MORRISON> %<SM:PDF_PARAM:MORRISON>
        Meal_Type = input('What type of breakfast would you like? (1 = Sweet, 2 = Savory) '); %<SM:PDF_PARAM:MORRISON>
        while isempty(Meal_Type)==1 || Meal_Type < 1 || Meal_Type > 2 || mod(Meal_Type,1) ~= 0
            Meal_Type = input('ERROR. What type of breakfast would you like? (1 = Sweet, 2 = Savory) ');
        end

    elseif Meal == 2
        Meal_Type = input('What type of protein would you like? (1 = Chicken, 2 = Beef, 3 = Shrimp, 4 = Tofu, 5 = Random) '); %<SM:RANDOM:MORRISON>
        while isempty(Meal_Type)==1 || Meal_Type < 1 || Meal_Type > 5 || mod(Meal_Type,1) ~= 0
            Meal_Type = (input('ERROR. What type of protein would you like? (1 = Chicken, 2 = Beef, 3 = Shrimp, 4 = Tofu, 5 = Random) '));
        end


    elseif Meal == 3
        Meal_Type = input('What type of protein would you like? (1 = Chicken, 2 = Steak, 3 = Fish, 4 = Tofu, 5 = Random) '); %<SM:RANDOM:MORRISON>
        while isempty(Meal_Type)==1 || Meal_Type < 1 || Meal_Type > 5 || mod(Meal_Type,1) ~= 0
            Meal_Type = str2double(input('ERROR. What type of protein would you like? (1 = Chicken, 2 = Steak, 3 = Fish, 4 = Tofu, 5 = Random) '));%<SM:RANDOM:MORRISON>
        end


    elseif Meal == 4
        Meal_Type = input('What type of dessert would you like? (1 = Fruit Based, 2 = Chocolate based, 3 = Unique) ');
        while isempty(Meal_Type)==1 || Meal_Type < 1 || Meal_Type > 3 || mod(Meal_Type,1) ~= 0
            Meal_Type = input('ERROR. What type of dessert would you like? (1 = Fruit Based, 2 = Chocolate based, 3 = Unique) ');
        end

    end

    if Meal_Type == 5
        Meal_Type = round(rand*3+1); %<SM:RANDOM:MORRISON>
    end

    Time = input('How much time do you have? (3 = 0-15 min, 4 = 15-30 min, 5 = 30-60 min) ');
    while isempty(Time)==1 || Time < 3 || Time > 5 || mod(Time,1)~=0
        Time = input('ERROR. How much time do you have? (3 = 0-15 min, 4 = 15-30 min, 5 = 30-60 min) ');
    end

    % Outputs

    fprintf('\nBased on your inputs, here are the Macros (Calories, Protein, Carbs and Fat) in grams.\n\n')

    [RowRec] = RecipeSearch(Recipes,Meal,Meal_Type); %<SM:PDF:MORRISON>
      
    


    url= char(RowRec{1,Time});  %<SM:VIEW:MORRISON> %This is used to make the links in excel into characters to be used to upload in the web.
    web(url)  %<SM:VIEW:MORRISON> % Citation 1

    if Time == 3
        Macros = [cell2mat(RowRec(6:9))];  %<SM:PDF_RETURN:MORRISON>  %This is used to separate the macros to make a bar chart.

        bar(Macros) %<SM:PLOT:MORRISON>
        set(gca,'xticklabel',{'Calories','Protein','Carbs','Fat'}); %Use matlab search label y axis citation
        text(1:length(Macros),Macros,num2str(Macros'),'vert','bottom','horiz','center'); % Citation 2
        [MIN,~] = min(Macros);
        [MAX,~] = max(Macros);
        ylim([(MIN-2) (MAX+22)])

        title('Recipe Macros in Grams')

    elseif Time == 4
        Macros = [cell2mat(RowRec(10:13))];

        bar(Macros) %<SM:PLOT:MORRISON>
        set(gca,'xticklabel',{'Calories','Protein','Carbs','Fat'}); 
        text(1:length(Macros),Macros,num2str(Macros'),'vert','bottom','horiz','center'); % Citation 2
        [MIN,~] = min(Macros);
        [MAX,~] = max(Macros);
        ylim([(MIN-2) (MAX+22)])

        title('Recipe Macros in Grams')


    elseif Time == 5
        Macros = [cell2mat(RowRec(14:17))];

        bar(Macros) %<SM:PLOT:MORRISON> %<SM:REF:MORRISON>
        set(gca,'xticklabel',{'Calories','Protein','Carbs','Fat'}); 
        text(1:length(Macros),Macros,num2str(Macros'),'vert','bottom','horiz','center'); % Citation 2
        [MIN,~] = min(Macros);
        [MAX,~] = max(Macros);
        ylim([(MIN-2) (MAX+22)])

        title('Recipe Macros in Grams')

    end


    Calories = [Macros(1)];
    Protein = [Macros(2)];
    Carbs = [Macros(3)];
    Fat = [Macros(4)];


    T= table(Calories,Protein,Carbs, Fat);
    disp(T) %<SM:VIEW:MORRISON>


    Run_again = (input('Do you want to run the program code again? (Yes or No) ','s')); %<SM:STRING:MORRISON>
    while isempty (Run_again) ==1 || isnan(str2double(Run_again))==0 || strcmpi(Run_again,'yes')==0 && strcmpi(Run_again,'no')==0
        Run_again = (input('Error. Do you want to run the program code again? (Yes or No) ','s'));
    end

end
fprintf('Thank you for using this program!\n')

