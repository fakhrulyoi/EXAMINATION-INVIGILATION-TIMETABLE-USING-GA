import random
import pandas as pd

# Initialize Population
def initialize_population(data_exam, data_invigilator, contoh_jadual, population_size=100):
    # Ensure invigilator data is clean
    data_invigilator = data_invigilator.dropna(subset=['Nama'])
    pensyarah_kanan = data_invigilator[data_invigilator['Jawatan'].str.contains('^PENSYARAH KANAN$', case=False, na=False)]

    if pensyarah_kanan.empty:
        raise ValueError("No eligible Pensyarah Kanan found for Ketua assignment!")

    population = []
    invigilation_counts = {name: 0 for name in data_invigilator['Nama']}

    def assign_random_venue():
        return random.choice(['DEWAN AKADEMIK', 'DEWAN LESTARI'])

    def classify_invigilator(invigilator_name, is_leader=False):
        invigilator = data_invigilator[data_invigilator['Nama'] == invigilator_name]
        if invigilator.empty:
            raise ValueError(f"Invigilator {invigilator_name} not found in data_invigilator!")
        invigilator = invigilator.iloc[0]
        if is_leader:
            return f"{invigilator_name} (K)"
        elif 'PENSYARAH' in invigilator['Jawatan']:
            return f"{invigilator_name} (L)"
        else:
            return f"{invigilator_name} (S)"

    for _ in range(population_size):
        schedule = []
        for _, row in data_exam.iterrows():
            exam_details = {
                'Tarikh': row['Tarikh'],
                'Hari': row['Hari'],
                'Waktu': row['Waktu'],
                'Kod Kursus': row['Kod Kursus'],
                'Masa': f"{row['Masa Mula']} - {row['Masa Tamat']}",
                'Bilangan Pelajar': row['Jumlah Pelajar'],
                'Tempat': assign_random_venue()
            }

            lecturers = contoh_jadual[contoh_jadual['Kod Kursus'] == row['Kod Kursus']]['Pensyarah'].unique()
            exam_details['Lecturer'] = ', '.join(lecturers) if len(lecturers) > 0 else 'Unknown'

            # Assign Ketua
            eligible_leaders = sorted(
                pensyarah_kanan['Nama'],
                key=lambda x: invigilation_counts[x]
            )
            leader = eligible_leaders[0]
            invigilation_counts[leader] += 1
            leader_label = classify_invigilator(leader, is_leader=True)

            # Assign Additional Invigilators
            required_invigilators = (row['Jumlah Pelajar'] + 29) // 30
            additional_invigilators = []
            eligible_invigilators = sorted(
                [name for name in data_invigilator['Nama'] if name not in lecturers],
                key=lambda x: invigilation_counts[x]
            )
            for _ in range(required_invigilators - 1):
                if eligible_invigilators:
                    new_invigilator = eligible_invigilators.pop(0)
                    additional_invigilators.append(new_invigilator)
                    invigilation_counts[new_invigilator] += 1

            invigilators = [leader_label] + [classify_invigilator(inv) for inv in additional_invigilators]
            exam_details['Invigilators'] = invigilators
            schedule.append(exam_details)

        population.append(schedule)

    return population

# Fitness Function
def calculate_fitness(schedule, data_exam, data_invigilator):
    fitness_score = 0
    invigilation_count = {name.lower(): 0 for name in data_invigilator['Nama']}

    # Get list of "Pensyarah Kanan"
    pensyarah_kanan = set(
        data_invigilator[data_invigilator['Jawatan'].str.contains('PENSYARAH KANAN', case=False, na=False)]['Nama'].str.lower()
    )

    for exam in schedule:
        # Constraint 1: Ketua must be Pensyarah Kanan
        ketua = exam['Invigilators'][0]
        ketua_name = ketua.split('(')[0].strip().lower()
        if ketua_name not in pensyarah_kanan:
            fitness_score += 2

        # Constraint 2: No lecturer should invigilate their own exam
        lecturers = [lecturer.strip().lower() for lecturer in exam['Lecturer'].split(',')]
        for invigilator in exam['Invigilators']:
            invigilator_name = invigilator.split('(')[0].strip().lower()
            if invigilator_name in lecturers:
                fitness_score += 2

        # Constraint 3: No male invigilators on Friday
        if exam['Hari'].strip().lower() == 'jumaat':
            for invigilator in exam['Invigilators']:
                invigilator_name = invigilator.split('(')[0].strip().lower()
                invigilator_record = data_invigilator[data_invigilator['Nama'].str.lower() == invigilator_name]
                if not invigilator_record.empty:
                    gender = invigilator_record['Jantina'].values[0].strip().lower()
                    if gender == 'male':
                        fitness_score += 1

        # Track invigilation counts
        for invigilator in exam['Invigilators']:
            invigilator_name = invigilator.split('(')[0].strip().lower()
            if invigilator_name in invigilation_count:
                invigilation_count[invigilator_name] += 1

    # Check invigilation count limits
    for name, count in invigilation_count.items():
        role = data_invigilator[data_invigilator['Nama'].str.lower() == name]['Jawatan'].values[0].lower()
        max_limit = 2 if 'pensyarah' in role else 1
        if count > max_limit:
            fitness_score += count - max_limit

    return fitness_score

# Parent Selection
def select_parents(population, fitness_scores, tournament_size=3):
    tournament = random.sample(range(len(population)), tournament_size)
    tournament_fitness = [fitness_scores[i] for i in tournament]
    return tournament[tournament_fitness.index(min(tournament_fitness))]

# Crossover
def perform_crossover(parent1, parent2):
    if not parent1 or not parent2:
        return parent1 if parent1 else parent2
    
    crossover_point = random.randint(1, len(parent1) - 1)
    child = parent1[:crossover_point] + parent2[crossover_point:]
    return child

# Mutation
def perform_mutation(schedule, data_invigilator, mutation_rate=0.0):
    if random.random() < mutation_rate:
        exam_idx = random.randint(0, len(schedule) - 1)
        pensyarah_kanan = data_invigilator[
            data_invigilator['Jawatan'].str.contains('PENSYARAH KANAN', case=False, na=False)
        ]
        if not pensyarah_kanan.empty:
            new_leader = random.choice(pensyarah_kanan['Nama'].tolist())
            schedule[exam_idx]['Invigilators'][0] = f"{new_leader} (K)"
    return schedule

# Create New Generation
def create_new_generation(population, data_exam, data_invigilator, elite_size=20):
    population_size = len(population)
    fitness_scores = [calculate_fitness(schedule, data_exam, data_invigilator) for schedule in population]
    
    # Sort population by fitness
    sorted_population = [x for _, x in sorted(zip(fitness_scores, population), key=lambda pair: pair[0])]
    
    new_population = []
    
    # Keep elite solutions
    new_population.extend(sorted_population[:elite_size])
    
    # Generate rest of population through crossover and mutation
    while len(new_population) < population_size:
        parent1_idx = select_parents(population, fitness_scores)
        parent2_idx = select_parents(population, fitness_scores)
        
        child = perform_crossover(population[parent1_idx], population[parent2_idx])
        child = perform_mutation(child, data_invigilator)
        
        new_population.append(child)
    
    return new_population

# Check Constraints
def check_constraints(schedule, data_exam, data_invigilator):
    violations = {
        'Ketua Not Pensyarah Kanan': {'count': 0, 'exams': []},
        'Lecturer Invigilating Own Exam': {'count': 0, 'exams': []},
        'Male Invigilator on Friday': {'count': 0, 'exams': []},
        'Exceeded Invigilation Limit': {'count': 0, 'exams': []},
        'Insufficient Invigilators': {'count': 0, 'exams': []}
    }
    
    pensyarah_kanan = set(
        data_invigilator[data_invigilator['Jawatan'].str.contains('PENSYARAH KANAN', case=False, na=False)]['Nama'].str.lower()
    )

    # Track invigilation counts for each person
    invigilation_count = {name.lower(): 0 for name in data_invigilator['Nama']}

    for exam in schedule:
        # Check Ketua constraint
        ketua = exam['Invigilators'][0]
        ketua_name = ketua.split('(')[0].strip().lower()
        if ketua_name not in pensyarah_kanan:
            violations['Ketua Not Pensyarah Kanan']['count'] += 2
            violations['Ketua Not Pensyarah Kanan']['exams'].append(exam)

        # Check lecturer invigilating own exam
        lecturers = [lecturer.strip().lower() for lecturer in exam['Lecturer'].split(',')]
        for invigilator in exam['Invigilators']:
            invigilator_name = invigilator.split('(')[0].strip().lower()
            if invigilator_name in lecturers:
                violations['Lecturer Invigilating Own Exam']['count'] += 2
                violations['Lecturer Invigilating Own Exam']['exams'].append(exam)

        # Check male invigilators on Friday
        if exam['Hari'].strip().lower() == 'jumaat':
            for invigilator in exam['Invigilators']:
                invigilator_name = invigilator.split('(')[0].strip().lower()
                invigilator_record = data_invigilator[data_invigilator['Nama'].str.lower() == invigilator_name]
                if not invigilator_record.empty:
                    gender = invigilator_record['Jantina'].values[0].strip().lower()
                    if gender == 'male':
                        violations['Male Invigilator on Friday']['count'] += 1
                        violations['Male Invigilator on Friday']['exams'].append(exam)

        # Check number of invigilators
        required_invigilators = (exam['Bilangan Pelajar'] + 29) // 30
        if len(exam['Invigilators']) < required_invigilators:
            violations['Insufficient Invigilators']['count'] += 2
            violations['Insufficient Invigilators']['exams'].append(exam)

        # Track invigilation counts
        for invigilator in exam['Invigilators']:
            invigilator_name = invigilator.split('(')[0].strip().lower()
            if invigilator_name in invigilation_count:
                invigilation_count[invigilator_name] += 1

    # Check invigilation count limits for each person
    for name, count in invigilation_count.items():
        role = data_invigilator[data_invigilator['Nama'].str.lower() == name]['Jawatan'].values[0].lower()
        max_limit = 2 if 'pensyarah' in role else 1
        if count > max_limit:
            violations['Exceeded Invigilation Limit']['count'] += (count - max_limit)
           

    return violations
# Main Genetic Algorithm
def genetic_algorithm(population, data_exam, data_invigilator, num_generations=30, target_fitness=0):
    best_schedule = None
    best_fitness = float('inf')
    best_violations = None
    generations_without_improvement = 0
    max_generations_without_improvement = 10

    for generation in range(num_generations):
        # Create new generation
        population = create_new_generation(population, data_exam, data_invigilator)
        
        # Evaluate population
        fitness_scores = [calculate_fitness(schedule, data_exam, data_invigilator) for schedule in population]
        current_best_idx = fitness_scores.index(min(fitness_scores))
        current_best_fitness = fitness_scores[current_best_idx]
        
        if current_best_fitness < best_fitness:
            best_fitness = current_best_fitness
            best_schedule = population[current_best_idx]
            best_violations = check_constraints(best_schedule, data_exam, data_invigilator)
            generations_without_improvement = 0
        else:
            generations_without_improvement += 1
        
        if best_fitness <= target_fitness or generations_without_improvement >= max_generations_without_improvement:
            break

    if best_violations is None and best_schedule is not None:
        best_violations = check_constraints(best_schedule, data_exam, data_invigilator)

        

    return best_schedule, best_fitness, best_violations
