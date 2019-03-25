import random
import pandas
import xlrd
import numpy
import time

"""Main function"""
def main():
    start_time = time.process_time()
    """define variables"""
    sheet = ['Mechanical']
    sheet_index = 0
    total = []
    iterations = 100
    iteration_count = 0
    population_size = 32

    """Read input data using file handling functions"""
    students = inputStudentFileHandling(sheet, sheet_index)
    projects = inputProjectFileHandling(sheet, sheet_index)
    
    """Initilise population"""
    allSolutions = initialisationLoop(students, projects, population_size, sheet, sheet_index)

    """Return students top 3 preferences"""
    choices = top3Preferences(sheet, sheet_index)

    """Return highest performing 4 solutions"""
    topIndex = fitnessFunction(sheet, sheet_index, allSolutions, population_size, choices)

    """Create 8 child solutions from highest performing parents"""
    allSolutions = crossoverMaster(topIndex, allSolutions, students)

    """Loop for specified number of iterations"""
    geneticLoop(sheet, sheet_index, population_size, iterations, allSolutions, choices, students)
    
    print (time.process_time() - start_time, "seconds")
    return

"""Function handling input student data"""
def inputStudentFileHandling(sheet, sheet_index):

    """Handles student data"""
    
    inputStudents = pandas.read_excel('StudentData.xlsx', sheet[sheet_index]) #read excel input data file

    """Imports ndarray of student codes"""
    raw_codes = inputStudents['Student Code'].values #assign imported data to a variable

    """Defines empty list for ndarrays of preferences to be imported into"""
    raw_choices = [] 

    """Imports preference data as listv of ndarrays for each preference set"""
    raw_choices.append(inputStudents['Project Code 1'].values)
    raw_choices.append(inputStudents['Project Code 2'].values)
    raw_choices.append(inputStudents['Project Code 3'].values)
    raw_choices.append(inputStudents['Project Code 4'].values)
    raw_choices.append(inputStudents['Project Code 5'].values)
    raw_choices.append(inputStudents['Project Code 6'].values)
    raw_choices.append(inputStudents['Project Code 7'].values)
    raw_choices.append(inputStudents['Project Code 8'].values)

    """Imports ndarray of student grades"""
    raw_grades = inputStudents['GPA'].values

    """creates empty lists to be populated with formatted data"""
    codes = [] 
    choices = [[],[],[],[],[],[],[],[]]
    grades = []
    
    """Formats raw student code data by converting ndarray into list"""
    for z in raw_codes: #creates list of student codes
        codes.append(z)

    """Turns ndarray object into list of lists"""
    count = 0 
    for z in raw_choices: #loops through the arrays for preferences
        for x in z: #loops throught the individial preferences within each array
            choices[count].append(float(x)) #appends the preference to the choices list as a float
        count = count + 1 #increments counter to point at the next preference list

    """Tidies up variables"""
    raw_choices = choices 
    choices = [] 

    """Creates a list of lists of the correct length for the amount of students in the dataset"""
    count = 0
    while count < len(codes): #runs until the total number of students in the dataset has been reached by count variable
        choices.append([])
        count = count + 1

    """Separates raw_choices list of lists into individual lists of preferences for each student"""
    for x in raw_choices: #loops through each preference list (all first preferences, second preferences... etc)
        count = 0
        for z in x: #for each individual preference
            choices[count].append(z) #appends preference list for each individual student with their preference
            count = count + 1  #increments count to look at the next student's preference list          
            
    """Formats raw grade data by converting ndarray into list"""
    for z in raw_grades: #creates list of student grades
        grades.append(z)

    """Format all imported data into lists of list"""
    students = [codes, choices, grades]

    return students

"""Function handling input project data"""
def inputProjectFileHandling(sheet, sheet_index):

    """Handles project data"""

    """Read excel file and assign sheet to a variable"""
    inputProjects = pandas.read_excel('ProjectData.xlsx', sheet[sheet_index]) 
    
    """Assign ndarray of project codes to variable"""
    raw_project_numbers = inputProjects['All Projects'].values

    """Create empty list for project codes to be appended to"""
    projects = []

    """Formats raw data by converting ndarray into list"""
    for z in raw_project_numbers:
        projects.append(z)

    return projects





def initialisePopulation(students, projects, matchedStudent, matchedProject, matchedPairs):  
    """Select random student and random project"""
    x = random.choice(students[0])
    y = random.choice(projects)
    
    """Append selections to lists"""
    matchedStudent.append(x)
    matchedProject.append(y)

    """Get student index for removal"""
    student_index = students[0].index(x)
    """Delete from previous dataset"""
    projects.remove(y)
    students[0].remove(students[0][student_index])
    students[1].remove(students[1][student_index])
    students[2].remove(students[2][student_index])

    """Return output"""
    return matchedPairs





"""Loop to create all populations"""

def initialisationLoop(students, projects, population_size, sheet, sheet_index):
    """Define count for population loop"""
    count = 0
    
    """Define list to hold completed population"""
    allSolutions = []
    
    """Loop creating as many solutions as the desired population size"""
    while count < population_size:
        """Define dictionary for student-project matches"""
        matchedStudent = []
        matchedProject = []
        matchedPairs = {'Student Code':matchedStudent, 'Project Code':matchedProject}


        """Read input data using file handling functions"""
        students = inputStudentFileHandling(sheet, sheet_index)
        projects = inputProjectFileHandling(sheet, sheet_index)
        
        """Define count for project assignment loop"""
        count2 = 0

        totalStudents = len(students[0])
        
        """Loop for as many students as there are in the dataset"""
        while count2 < totalStudents:
            initialisePopulation(students, projects, matchedStudent, matchedProject, matchedPairs)
            count2 = count2 + 1
            

        """Append solution to list"""
        allSolutions.append(matchedPairs)

        count = count + 1
        
    """Return list of dictionaries holding each solution"""
    return allSolutions






"""Function to retrieve a student's top 3 preferences"""
"""This operates almost identically to the input file functions, but just imports 3 preferences rather than all 8"""
def top3Preferences(sheet, sheet_index):
    inputStudents = pandas.read_excel('StudentData.xlsx', sheet[sheet_index]) #read excel input data file
    
    raw_codes = inputStudents['Student Code'].values #assign imported data to a variable
    codes = []

    for z in raw_codes: #creates list of student codes
        codes.append(z)

    raw_choices = []

    """Gets top 3 preferences from input file"""
    
    raw_choices.append(inputStudents['Project Code 1'].values)
    raw_choices.append(inputStudents['Project Code 2'].values)
    raw_choices.append(inputStudents['Project Code 3'].values)

    choices = [[],[],[]] #define empty list of lists to house reformatted ndarray objects

    count = 0 
    for z in raw_choices: #turns ndarray object into list of lists
        for x in z:
            choices[count].append(float(x))
        count = count + 1

    raw_choices = choices #puts list of lists from choices into raw_choices so choices can be reused
    choices = [] #resets choices list variable

    count = 0
    while count < len(codes): #creates empty choices list of lists of correct length for student preferences
        choices.append([])
        count = count + 1


    for x in raw_choices: #formats student data to individual lists of preferences inside choices list
        count = 0
        for z in x:
            choices[count].append(z)
            count = count + 1

    return choices





        
"""Returns a list containing the fitness function of each solution"""
def fitnessCalculator(sheet, sheet_index, studentProjectPairs, totalPercentages, choices):
    """Variable holding the output"""
    
    """Loop to establish number of top 3 choices verses number of lower choices"""
    highChoice = 0
    lowChoice = 0
    count = 0
    for x in studentProjectPairs: #points at the project the student has been matched with
        for z in choices[count]: #points at the list of top 3 preferences of the student
            if x[1] == z: #if there is a match between the project they were assigned and one of their top 3 preferences
                highChoice = highChoice + 1 #increment the highChoice variable
                count = count+1
                break
        else: #if there is no match
            lowChoice = lowChoice + 1 #increment the lowChoice variable
            count = count + 1 #increment the count variable to look at the next student
            
        

    """Calculate percentage"""
    total = highChoice + lowChoice
    
    highChoicePercentage = (highChoice/total)*100

    totalPercentages.append(highChoicePercentage)
    
    """Return list of fitness functions"""
    return totalPercentages
    
 
        


"""Function to extract the solutions with the highest fitness function and their indexes"""
def highestFitness(totalPercentages, population_size):
    """Define variables"""
    topPerformers = []
    topIndex = []
    count = 0
    total = population_size/2


    
    """While loop - loops for half the number of solutions"""
    while count < total:
        refVar = 0
        """For loop - iterates over list containing solution fitnesses and assignes the highest scoring fitness to a reference variable"""
        for x in totalPercentages:
            if x > refVar:
                refVar = x
                
        """Appends topPerformers list with the highest performing solution's fitness"""        
        topPerformers.append(refVar)
        """Appends topIndex list with the index of the highest performing solution"""
        topIndex.append(totalPercentages.index(refVar))
        """Saves that index to a variable"""
        indexVar = totalPercentages.index(refVar)
        """Removes the highest performer from the list, so that it does not get selected again in the next iteration"""
        totalPercentages.remove(refVar)
        """Replaces it with a 0 to ensure correct indexing is maintained"""
        totalPercentages.insert(indexVar, 0)
        count = count + 1

    return topIndex
    






def fitnessFunction(sheet, sheet_index, allSolutions, population_size, choices):
    """Reformat and sort data for fitness calculation"""
    totalPercentages = []
    for x in allSolutions:
        """Define variables to hold data"""
        a = 0 #will hold the student codes
        b = 0 #will hold the project codes
        count = 0
        tempdict = {} #dict to hold the values for sorting

        """Unpacks the student and project codes for the solution"""
        a = x.get('Student Code')
        b = x.get('Project Code')

        """Iterate through each student and project code and save to dictionary as a key value pair"""
        while count < len(a):
            tempdict[a[count]] = b[count]
            count = count + 1
            
        """Sorts through the key value pair so data is sorted by student code"""
        """Data is outputted as a list of tuples in the format [(student, project)]"""
        studentProjectPairs = sorted(tempdict.items())

        """Outputs a list containing the fitness of each member of the population"""
        """The list increases in size with each iteration"""
        totalPercentages = fitnessCalculator(sheet, sheet_index, studentProjectPairs, totalPercentages, choices)

    print(totalPercentages)
    topIndex = highestFitness(totalPercentages, population_size)
    """Returns the highest performing functions and their indexes"""
    return topIndex
        

"""Function performing crossover and outputting 8 new child solutions"""
def crossoverMaster(topIndex, allSolutions, students): #remember that allSolutions is a list holding 32 dictionaries, 1 for each solution
    """Call slicing function"""
    slicedListMaster = crossoverListSlicing(topIndex, allSolutions)

    """Call mutation function"""
    allSolutions = mutation(slicedListMaster, students)

    """Call child generation function"""
    allSolutions = crossoverSolutions(slicedListMaster, allSolutions)

    return allSolutions

"""Separate solutions into sections for crossover"""
def crossoverListSlicing(topIndex, allSolutions):
    """Define variables that randomly hold a different index"""
    """Indexes are then removed from topIndex variable so they are not repeatedly assigned"""
    s1 = random.choice(topIndex)
    topIndex.remove(s1)
    s2 = random.choice(topIndex)
    topIndex.remove(s2)
    s3 = random.choice(topIndex)
    topIndex.remove(s3)
    s4 = random.choice(topIndex)
    topIndex.remove(s4)
    s5 = random.choice(topIndex)
    topIndex.remove(s5)
    s6 = random.choice(topIndex)
    topIndex.remove(s6)
    s7 = random.choice(topIndex)
    topIndex.remove(s7)
    s8 = random.choice(topIndex)
    topIndex.remove(s8)


    """Assign solution dictionaries to variables"""
    s1 = allSolutions[s1]
    s2 = allSolutions[s2]
    s3 = allSolutions[s3]
    s4 = allSolutions[s4]
    s5 = allSolutions[s5]
    s6 = allSolutions[s6]
    s7 = allSolutions[s7]
    s8 = allSolutions[s8]

    """Slice up solutions"""
    """List to hold each half of the student and project codes"""
    """Naming convention is (List - solution number - section of solution - student or project)"""
    """Projects and students for solution 1"""
    Ls11P = s1.get('Project Code')[:int(len(s1.get('Project Code'))/2)] #first half of projects from solution 1
    Ls12P = s1.get('Project Code')[int(len(s1.get('Project Code'))/2):] #second half of projects from solution 1
    Ls11S = s1.get('Student Code')[:int(len(s1.get('Student Code'))/2)] #first half of students from solution 1
    Ls12S = s1.get('Student Code')[int(len(s1.get('Student Code'))/2):] #second half of students from solution 1
    
    """Projects and students for solution 2"""
    Ls21P = s2.get('Project Code')[:int(len(s2.get('Project Code'))/2)] #first half of projects from solution 2
    Ls22P = s2.get('Project Code')[int(len(s2.get('Project Code'))/2):] #second half of projects from solution 2
    Ls21S = s2.get('Student Code')[:int(len(s2.get('Student Code'))/2)] #first half of students from solution 2
    Ls22S = s2.get('Student Code')[int(len(s2.get('Student Code'))/2):] #second half of students from solution 2

    """Projects and students for solution 3"""
    Ls31P = s3.get('Project Code')[:int(len(s3.get('Project Code'))/2)] #first half of projects from solution 3
    Ls32P = s3.get('Project Code')[int(len(s3.get('Project Code'))/2):] #second half of projects from solution 3
    Ls31S = s3.get('Student Code')[:int(len(s3.get('Student Code'))/2)] #first half of students from solution 3
    Ls32S = s3.get('Student Code')[int(len(s3.get('Student Code'))/2):] #second half of students from solution 3
    
    """Projects and students for solution 4"""
    Ls41P = s4.get('Project Code')[:int(len(s4.get('Project Code'))/2)] #first half of projects from solution 4
    Ls42P = s4.get('Project Code')[int(len(s4.get('Project Code'))/2):] #second half of projects from solution 4
    Ls41S = s4.get('Student Code')[:int(len(s4.get('Student Code'))/2)] #first half of students from solution 4
    Ls42S = s4.get('Student Code')[int(len(s4.get('Student Code'))/2):] #second half of students from solution 4

    """Projects and students for solution 5"""
    Ls51P = s5.get('Project Code')[:int(len(s5.get('Project Code'))/2)] #first half of projects from solution 5
    Ls52P = s5.get('Project Code')[int(len(s5.get('Project Code'))/2):] #second half of projects from solution 5
    Ls51S = s5.get('Student Code')[:int(len(s5.get('Student Code'))/2)] #first half of students from solution 5
    Ls52S = s5.get('Student Code')[int(len(s5.get('Student Code'))/2):] #second half of students from solution 5
    
    """Projects and students for solution 6"""
    Ls61P = s6.get('Project Code')[:int(len(s6.get('Project Code'))/2)] #first half of projects from solution 6
    Ls62P = s6.get('Project Code')[int(len(s6.get('Project Code'))/2):] #second half of projects from solution 6
    Ls61S = s6.get('Student Code')[:int(len(s6.get('Student Code'))/2)] #first half of students from solution 6
    Ls62S = s6.get('Student Code')[int(len(s6.get('Student Code'))/2):] #second half of students from solution 6
    
    """Projects and students for solution 7"""
    Ls71P = s7.get('Project Code')[:int(len(s7.get('Project Code'))/2)] #first half of projects from solution 7
    Ls72P = s7.get('Project Code')[int(len(s7.get('Project Code'))/2):] #second half of projects from solution 7
    Ls71S = s7.get('Student Code')[:int(len(s7.get('Student Code'))/2)] #first half of students from solution 7
    Ls72S = s7.get('Student Code')[int(len(s7.get('Student Code'))/2):] #second half of students from solution 7
    
    """Projects and students for solution 8"""
    Ls81P = s8.get('Project Code')[:int(len(s8.get('Project Code'))/2)] #first half of projects from solution 8
    Ls82P = s8.get('Project Code')[int(len(s8.get('Project Code'))/2):] #second half of projects from solution 8
    Ls81S = s8.get('Student Code')[:int(len(s8.get('Student Code'))/2)] #first half of students from solution 8
    Ls82S = s8.get('Student Code')[int(len(s8.get('Student Code'))/2):] #second half of students from solution 8

    """List to hold sliced lists"""
    slicedStudentLists1 = [Ls11S, Ls21S, Ls31S, Ls41S, Ls51S, Ls61S, Ls71S, Ls81S]
    slicedStudentLists2 = [Ls12S, Ls22S, Ls32S, Ls42S, Ls52S, Ls62S, Ls72S, Ls82S]
    slicedProjectLists1 = [Ls11P, Ls21P, Ls31P, Ls41P, Ls51P, Ls61P, Ls71P, Ls81P]
    slicedProjectLists2 = [Ls12P, Ls22P, Ls32P, Ls42P, Ls52P, Ls62P, Ls72P, Ls82P]
    slicedListMaster = [slicedStudentLists1, slicedStudentLists2, slicedProjectLists1, slicedProjectLists2]
    return slicedListMaster


"""Generate child solutions"""
def crossoverSolutions(slicedListMaster, allSolutions):
    """Unpack master list"""
    slicedStudentLists1 = slicedListMaster[0]
    slicedStudentLists2 = slicedListMaster[1]
    slicedProjectLists1 = slicedListMaster[2]
    slicedProjectLists2 = slicedListMaster[3]
    """Loop to randomly generate 16 children"""
    count = 0
    allSolutions = []
    while count < 16:
        studentChild = 0
        projectChild = 0
        studentChild = random.choice(slicedStudentLists1) + random.choice(slicedStudentLists2)
        projectChild = random.choice(slicedProjectLists1) + random.choice(slicedProjectLists2)
        crossoverDict = {'Student Code':studentChild, 'Project Code':projectChild}
        allSolutions.append(crossoverDict)
        count = count + 1
    count = 0
    while count < 16:
        studentChild = 0
        projectChild = 0
        studentChild = random.choice(slicedStudentLists2) + random.choice(slicedStudentLists1)
        projectChild = random.choice(slicedProjectLists2) + random.choice(slicedProjectLists1)
        crossoverDict = {'Student Code':studentChild, 'Project Code':projectChild}
        allSolutions.append(crossoverDict)
        count = count + 1
    
    return allSolutions

"""Mutation function to ensure the avoidance of a local maximum"""
def mutation(slicedListMaster, students): 
    """Define mutation rate at 10% of the population"""
    mutationRate = int(len(students[0])/10)
    """Variables for data handling"""
    studentLists = [slicedListMaster[0], slicedListMaster[1]]
    projectLists = [slicedListMaster[2], slicedListMaster[3]]
    """Randomly select 10% of each solution for mutation"""
    """The first loop is for each half of the student codes"""
    studentSampleIndexMaster = []
    for x in studentLists:
        """The second loop looks at each indiviual half student code set"""
        for z in x:
            """Randomly selects student codes at the mutation rate"""
            studentSample = random.sample(z, mutationRate)
            studentSampleIndex = []
            """This loop returns the indexes of the selected students"""
            for y in studentSample:
                studentSampleIndex.append(z.index(y))
            """Append index list to master list for later handling"""
            studentSampleIndexMaster.append(studentSampleIndex)
    """The output is a list of lists containing sampled student indexes for each half solution set"""


    """Use these indexes to extract and shuffle the associated projects assigned to those students"""
    projectSamplesMaster = [] #this will be a list of lists holding the sampled projects
    count = 0 #at count = 0, the first half lists are iterated through. at count = 1, the second half lists are iterated through
    while count < 2:

        if count == 0:
            for halfList in projectLists[0]: #first half of projects assigned
                projectSamples = []
                for studentIndex in studentSampleIndexMaster[projectLists[0].index(halfList)]: #only look at the same student list as the project list being looked at
                    projectSamples.append(halfList[studentIndex]) #append projectSamples with project at same index as the student sampled
                random.shuffle(projectSamples) #shuffle the list
                projectSamplesMaster.append(projectSamples) #append the list to the master list of lists
        """projectSamplesMaster should contain 8 lists each with 12 elements"""
        if count == 1:

            for halfList in projectLists[1]: #second half of projects assigned
                projectSamples = []
                for studentIndex in studentSampleIndexMaster[projectLists[1].index(halfList)]:
                    projectSamples.append(halfList[studentIndex])
                random.shuffle(projectSamples)
                projectSamplesMaster.append(projectSamples)
        """projectSamplesMaster should contain 16 lists each with 12 elements"""
        
            
                
        count = count + 1


    count = 0

    """Reassignes projects at sampled indexes"""
    while count < 2:

        if count == 0:
            for halfList in projectLists[0]: #first half of assinged projects
                count2 = 0
                for sampledIndex in studentSampleIndexMaster[projectLists[0].index(halfList)]: #only look at the same index list as the project list being looked at
                    randomChoice = random.choice(projectSamplesMaster[projectLists[0].index(halfList)])
                    projectLists[0][projectLists[0].index(halfList)][sampledIndex] = randomChoice #replace project at sampled index with project from sampled list 
                    projectSamplesMaster[projectLists[0].index(halfList)].remove(randomChoice)
       
                    
        if count == 1:
            for halfList in projectLists[1]:
                count2 = 0
                for sampledIndex in studentSampleIndexMaster[projectLists[1].index(halfList)]:
                    randomChoice = random.choice(projectSamplesMaster[projectLists[1].index(halfList) + 8])
                    projectLists[1][projectLists[1].index(halfList)][sampledIndex] = randomChoice
                    projectSamplesMaster[projectLists[1].index(halfList) + 8].remove(randomChoice)
                 
                                    
        count = count + 1

        
                                    

    slicedListMaster = [studentLists, projectLists]   
    return slicedListMaster


"""Function to run program as many times as there are specified iterations"""
def geneticLoop(sheet, sheet_index, population_size, iterations, allSolutions, choices, students):
    count = 0
    while count < iterations:
        
        """Run fitness calculation function"""
        topIndex = fitnessFunction(sheet, sheet_index, allSolutions, population_size, choices)

        """Run crossover function"""
        allSolutions = crossoverMaster(topIndex, allSolutions, students)

        count = count + 1
    return allSolutions
    
        
