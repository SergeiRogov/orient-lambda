from bs4 import BeautifulSoup
import boto3
import json
import re


class Runner:
    def __init__(self, name: str, course: str, place: int, bib: int, age_group: str, overall_time: str,
                 splits: list[tuple[str, int, int, int, int, int, int, str, str]], id: int):
        self.name = name
        self.course = course
        self.place = place
        self.bib = bib
        self.age_group = age_group
        self.overall_time = overall_time
        self.splits = splits
        self.id = id

    def to_json(self):
        # Return a dictionary representation of the object
        return {
            'name': self.name,
            'course': self.course,
            'place': self.place,
            'bib': self.bib,
            'age_group': self.age_group,
            'overall_time': self.overall_time,
            'splits': self.splits,
            'id': self.id
        }


s3 = boto3.client('s3')

control_point = re.compile(r'^\d{1,2}\(\d{2,3}\)\s*')  # control point "3(57)"
course_title = re.compile(r'\w+\s*\(\d{1,3}\)\s*')  # course title "Blue (16)"
split_time = re.compile(r'\d{1,2}:\d{2}')  # time "03:22"


def has_child_i_tag(tag):
    return tag.find('i', recursive=False) is not None


def convert_to_seconds(time_str):
    if time_str == '----':
        return float('inf')
    minutes, seconds = map(int, time_str.split(':'))
    return minutes * 60 + seconds


def convert_to_str(total_seconds):
    return f"{total_seconds // 60:02}:{total_seconds % 60:02}"


def split_key(obj):
    return obj.get('split', float('inf'))


def cumul_key(obj):
    return obj.get('cumulative', float('inf'))


def parse_raw_html(html_filename, html_content):

    soup = BeautifulSoup(html_content, 'html.parser')
    title = soup.find('title').text.strip()

    if html_filename == "Kalavasos30_09_23splits.html":
        # delete empty tags for specific case
        for x in soup.find_all():
            if len(x.get_text(strip=True)) == 0 and x.name not in ['font'] and len(x.find_all()) == 0:
                x.extract()
    else:
        # delete empty tags
        for x in soup.find_all():
            if len(x.get_text(strip=True)) == 0 and x.name not in ['font'] and len(x.find_all()) == 0 \
                    or has_child_i_tag(x):
                x.extract()

    # Names of courses
    courses = soup.find_all('b', string=course_title)
    courses = [re.sub(r"\s*\(\d+\)", '', title.text.strip()) for title in courses]

    # Find the table elements with results for all 3 courses (and header coming first)
    tables = soup.find_all('table')

    # Lists to store control points and runners for all courses
    all_courses_controls = []
    all_courses_runners = []
    # Iterating through each course (=table) and extracting splits
    for course_index, table in enumerate(tables[1:]):

        course_controls = []
        course_runners = []
        runner_splits = []

        place_num = 0
        bib = 0
        name = ''
        group = ''
        time = ''

        num_of_font_cells = 0
        info_row = None
        num_of_control_rows = None
        is_runner_extracted = False

        # Extract table's rows
        rows = table.find_all('tr')

        # Process each row
        for row_index, row in enumerate(rows):
            # Skip redundant row (with separated splits)
            if html_filename != "Kalavasos30_09_23splits.html":
                if info_row and num_of_control_rows and \
                        (info_row & 1) != (
                        row_index & 1) and row_index - info_row < num_of_control_rows * 2 - 1:
                    continue

            is_runner_found = False
            runner_info_start_index = None

            # Extract row's cells
            cells = row.find_all('td')

            if len(cells) == 0 or (is_runner_extracted and not cells[0].find('font')):
                continue

            # Process each cell
            for cell_index, cell in enumerate(cells):
                # Extracting course controls' sequence
                if cell.find(string=control_point) or cell.text.strip() == "F":
                    course_controls.append(cell.text.strip())
                    if cell.text.strip() == "F":
                        # Number of rows covering a course
                        # if it's a single row course (row_index = 0) - it still would be a 2 row course in splits
                        num_of_control_rows = row_index + 1 if row_index > 0 else 2

                # Finding first field related to runner (first <font> tag)
                elif not is_runner_found and cell.find('font'):
                    is_runner_found = True
                    is_runner_extracted = False
                    runner_splits = []

                    runner_info_start_index = cell_index
                    info_row = row_index

                    num_of_font_cells = 1
                    i = cell_index
                    while cells[i + 1].find('font'):
                        num_of_font_cells += 1
                        i += 1

                    # Recording info from this field
                    if num_of_font_cells == 5:
                        place_num = cell.text.strip()
                    else:
                        place_num = "<-->"
                        bib = cell.text.strip()

                # Recording runners' info from the 4 following fields
                # Recording splits from the same row
                elif is_runner_found and runner_info_start_index is not None:
                    if cell_index < num_of_font_cells:
                        if num_of_font_cells == 5:
                            if cell_index == runner_info_start_index + 1:
                                bib = cell.text.strip() if cell.text.strip() != '' else "<-->"
                            elif cell_index == runner_info_start_index + 2:
                                name = cell.text.strip() if cell.text.strip() != '' else "<-->"
                            elif cell_index == runner_info_start_index + 3:
                                group = cell.text.strip() if cell.text.strip() != '' else "<-->"
                            elif cell_index == runner_info_start_index + 4:
                                time = cell.text.strip() if cell.text.strip() != '' else "<-->"
                        else:
                            if cell_index == runner_info_start_index + 1:
                                name = cell.text.strip() if cell.text.strip() != '' else "<-->"
                            elif cell_index == runner_info_start_index + 2:
                                group = cell.text.strip() if cell.text.strip() != '' else "<-->"
                            elif cell_index == runner_info_start_index + 3:
                                time = cell.text.strip() if cell.text.strip() != '' else "<-->"
                    # splits
                    elif cell_index >= num_of_font_cells and len(runner_splits) < len(course_controls):
                        if cell.find(string=split_time):
                            runner_splits.append(cell.text.strip())
                            # if splits end on the same row they start - add a runner
                            if len(runner_splits) == len(course_controls):
                                runner_instance = Runner(name, courses[course_index], place_num, bib, group,
                                                         time,
                                                         runner_splits, len(course_runners))
                                course_runners.append(runner_instance.to_json())
                                is_runner_extracted = True
                                break
                        elif time == "DNF":
                            # Adding not fully completed runner info (if splits are not right)
                            runner_instance = Runner(name, courses[course_index], place_num, bib, group, time,
                                                     runner_splits, len(course_runners))
                            course_runners.append(runner_instance.to_json())
                            is_runner_extracted = True
                            break
                        else:
                            runner_splits.append("----")
                # Adding splits from other rows with same parity
                elif (info_row & 1) == (row_index & 1) and info_row != row_index:
                    if cell.find(string=split_time):
                        runner_splits.append(cell.text.strip())
                        if len(runner_splits) == len(course_controls):
                            # Adding runner info
                            runner_instance = Runner(name, courses[course_index], place_num, bib, group, time,
                                                     runner_splits, len(course_runners))
                            course_runners.append(runner_instance.to_json())
                            is_runner_extracted = True
                            break
                    elif time == "DNF":
                        # Adding not fully completed runner info (splits are not right)
                        if not is_runner_extracted:
                            runner_instance = Runner(name, courses[course_index], place_num, bib, group, time,
                                                     runner_splits, len(course_runners))
                            course_runners.append(runner_instance.to_json())
                            is_runner_extracted = True
                            break
                    else:
                        runner_splits.append("----")

        all_courses_controls.append(course_controls)
        all_courses_runners.append(course_runners)
    return title, courses, all_courses_controls, all_courses_runners


def calculate_splits_from_cumulative_timestamps(all_courses_runners):
    splits = []
    for course_index in range(len(all_courses_runners)):
        course_splits = []
        for runner_index in range(len(all_courses_runners[course_index])):
            runner_splits = []
            for control_index in range(len(all_courses_runners[course_index][runner_index]['splits']) - 1, 0,
                                       -1):
                current = convert_to_seconds(
                    all_courses_runners[course_index][runner_index]['splits'][control_index])
                previous = convert_to_seconds(
                    all_courses_runners[course_index][runner_index]['splits'][control_index - 1])
                split_time = current - previous
                all_courses_runners[course_index][runner_index]['splits'][
                    control_index] = f"{convert_to_str(split_time)}*{all_courses_runners[course_index][runner_index]['splits'][control_index]}"
                runner_splits.append(
                    {'split': split_time,
                     'cumulative': current,
                     'runner_index': runner_index,
                     'control_index': control_index})

            if len(all_courses_runners[course_index][runner_index]['splits']) > 0:
                # first split separately
                first_split = all_courses_runners[course_index][runner_index]['splits'][0]
                all_courses_runners[course_index][runner_index]['splits'][0] = f"{first_split}*{first_split}"
                runner_splits.append(
                    {'split': convert_to_seconds(first_split),
                     'cumulative': convert_to_seconds(first_split),
                     'runner_index': runner_index,
                     'control_index': 0})

            runner_splits.reverse()
            course_splits.append(runner_splits)
        max_row_length = max(len(row) for row in course_splits)

        # Add zeros to rows with fewer elements
        course_splits = [row + [{}] * (max_row_length - len(row)) for row in course_splits]

        splits.append(course_splits)
    return splits


def sort_runners_for_each_leg(splits):
    sorted_split_runners = []
    sorted_cumul_runners = []
    for course_index in range(len(splits)):
        transposed_runners = list(map(list, zip(*splits[course_index])))
        sorted_split_runners_course = []
        sorted_cumul_runners_course = []
        for control in transposed_runners:
            sorted_split = [sorted(control, key=split_key)]
            sorted_split_runners_course.append(list(map(list, zip(*sorted_split))))

            sorted_cumulative = [sorted(control, key=cumul_key)]
            sorted_cumul_runners_course.append(list(map(list, zip(*sorted_cumulative))))

        sorted_split_runners.append(sorted_split_runners_course)
        sorted_cumul_runners.append(sorted_cumul_runners_course)
    return sorted_split_runners, sorted_cumul_runners


def add_split_positions(sorted_split_runners, all_courses_runners):
    for course_index in range(len(sorted_split_runners)):
        for control_index in range(len(sorted_split_runners[course_index])):
            for runner_leg_place in range(len(sorted_split_runners[course_index][control_index])):
                if 'runner_index' in sorted_split_runners[course_index][control_index][runner_leg_place][0]:
                    runner_num = sorted_split_runners[course_index][control_index][runner_leg_place][0][
                        'runner_index']
                    cell = all_courses_runners[course_index][runner_num]['splits'][control_index]
                    split, cumul = map(str, cell.split('*'))
                    if runner_leg_place > 0 and \
                            sorted_split_runners[course_index][control_index][runner_leg_place][0]['split'] == \
                            sorted_split_runners[course_index][control_index][runner_leg_place - 1][0]['split']:
                        cur_place = all_courses_runners[course_index][prev_runner_num]['splits'][control_index][
                            1]
                        new_cell_tuple = (f"{split}({cur_place})*{cumul}", cur_place)
                    else:
                        cur_place = runner_leg_place + 1
                        new_cell_tuple = (f"{split}({cur_place})*{cumul}", cur_place)
                    prev_runner_num = runner_num
                    all_courses_runners[course_index][runner_num]['splits'][control_index] = new_cell_tuple


def add_cumul_positions(sorted_cumul_runners, all_courses_runners):
    for course_index in range(len(sorted_cumul_runners)):
        for control_index in range(len(sorted_cumul_runners[course_index])):
            for runner_cumul_place in range(len(sorted_cumul_runners[course_index][control_index])):
                if 'runner_index' in sorted_cumul_runners[course_index][control_index][runner_cumul_place][0]:
                    runner_num = sorted_cumul_runners[course_index][control_index][runner_cumul_place][0][
                        'runner_index']
                    cell = all_courses_runners[course_index][runner_num]['splits'][control_index]

                    if runner_cumul_place > 0 and \
                            sorted_cumul_runners[course_index][control_index][runner_cumul_place][0][
                                'cumulative'] == \
                            sorted_cumul_runners[course_index][control_index][runner_cumul_place - 1][0][
                                'cumulative']:
                        cur_place = all_courses_runners[course_index][prev_runner_num]['splits'][control_index][
                            2]
                        new_cell_tuple = (f"{cell[0]}({cur_place})", cell[1], cur_place)
                    else:
                        cur_place = runner_cumul_place + 1
                        new_cell_tuple = (f"{cell[0]}({cur_place})", cell[1], cur_place)

                    prev_runner_num = runner_num
                    all_courses_runners[course_index][runner_num]['splits'][control_index] = new_cell_tuple


def add_loss_info(sorted_split_runners, sorted_cumul_runners, splits, all_courses_runners):
    for course_index in range(len(all_courses_runners)):
        for runner_index in range(len(all_courses_runners[course_index])):
            for control_index in range(len(all_courses_runners[course_index][runner_index]['splits'])):
                best_split = sorted_split_runners[course_index][control_index][0][0]['split']
                best_cumul = sorted_cumul_runners[course_index][control_index][0][0]['cumulative']

                if len(sorted_split_runners[course_index][control_index]) > 1:
                    second_best_split = sorted_split_runners[course_index][control_index][1][0]['split']
                    second_best_cumul = sorted_cumul_runners[course_index][control_index][1][0]['cumulative']
                else:
                    second_best_split = best_split
                    second_best_cumul = best_cumul

                cell = all_courses_runners[course_index][runner_index]['splits'][control_index]

                split = splits[course_index][runner_index][control_index]['split']
                cumul = splits[course_index][runner_index][control_index]['cumulative']

                behind_best_split = "+" + convert_to_str(
                    split - best_split) + f' +{round(split / best_split * 100 - 100, 0):.0f}%' if split != best_split else "-" + convert_to_str(
                    second_best_split - split) + f' -{round(second_best_split / split * 100 - 100, 0):.0f}%'
                behind_best_cumul = "+" + convert_to_str(
                    cumul - best_cumul) + f' +{round(cumul / best_cumul * 100 - 100, 0):.0f}%' if cumul != best_cumul else "-" + convert_to_str(
                    second_best_cumul - cumul) + f' -{round(second_best_cumul / cumul * 100 - 100, 0):.0f}%'

                all_courses_runners[course_index][runner_index]['splits'][control_index] = (
                    cell[0],  # string to display
                    cell[1],  # place split
                    cell[2],  # place cumul
                    round(split / best_split * 100, 0),  # percent from best split
                    round(cumul / best_cumul * 100, 0),  # percent from best cumul
                    split,  # split in seconds
                    cumul,  # cumul in seconds,
                    str(behind_best_split),  # info to be displayed on hover (split)
                    str(behind_best_cumul)  # info to be displayed on hover (cumul)
                )


def fill_missing_splits_with_proportional_ratio(all_courses_runners):
    for course_index, runners in enumerate(all_courses_runners):
        for runner_index, runner in enumerate(runners):
            splits = runner['splits']
            missing_ranges = find_consecutive_missing_ranges(splits)

            for start, end in missing_ranges:
                if start == 0:
                    # Missing splits are at the beginning
                    all_courses_runners[course_index][runner_index]['splits'][start:end + 1] = estimate_start_gap(
                        runners, splits, end)
                elif end == len(splits) - 1:
                    # Missing splits are at the end
                    if runner['overall_time'] != "DNF":
                        all_courses_runners[course_index][runner_index]['splits'][-1] = runner['overall_time']
                        if start < end:
                            all_courses_runners[course_index][runner_index]['splits'][start:end] = estimate_middle_gap(
                                runners, splits, start, end - 1)
                else:
                    # Missing splits in the middle
                    all_courses_runners[course_index][runner_index]['splits'][start:end + 1] = estimate_middle_gap(
                        runners, splits, start, end)


def find_consecutive_missing_ranges(splits):
    missing_ranges = []
    start = None

    for i, split in enumerate(splits):
        if split in ["----", "00:00"]:
            if start is None:
                start = i
        elif start is not None:
            missing_ranges.append((start, i - 1))
            start = None

    if start is not None:
        missing_ranges.append((start, len(splits) - 1))

    return missing_ranges


def estimate_start_gap(runners, splits, end):
    next_split_time = convert_to_seconds(splits[end + 1])
    ratios = calculate_best_ratios(runners, 0, end + 1)
    return proportional_interpolation(ratios, next_split_time=next_split_time)


def estimate_middle_gap(runners, splits, start, end):
    if start > end:
        return
    prev_split_time = convert_to_seconds(splits[start - 1])
    next_split_time = convert_to_seconds(splits[end + 1])
    ratios = calculate_best_ratios(runners, start, end)
    return proportional_interpolation(ratios, prev_split_time=prev_split_time,
                                      next_split_time=next_split_time)


def calculate_best_ratios(runners, start, end):
    if start > end:
        return

    leg_times = []
    total_best_time = 0

    for i in range(start - 1, end + 1):
        if i < 0 or i + 1 >= len(runners[0]['splits']):  # Skip out-of-bounds indices
            continue

        best_time = float('inf')
        for runner in runners:
            splits = runner['splits']

            # Check that both splits at i and i+1 exist and are not missing
            if i + 1 < len(splits) and splits[i] != "----" and splits[i + 1] != "----":
                time_diff = convert_to_seconds(splits[i + 1]) - convert_to_seconds(splits[i])
                best_time = min(best_time, time_diff)

        # If a valid best time is found, add it to leg_times and accumulate total time
        if best_time != float('inf'):
            leg_times.append(best_time)
            total_best_time += best_time

    # Calculate ratios for each leg relative to the total best time
    return [leg_time / total_best_time for leg_time in leg_times] if total_best_time > 0 else []


def proportional_interpolation(ratios, prev_split_time=None, next_split_time=None):
    interpolated_splits = []

    if prev_split_time is not None and next_split_time is not None:
        # Calculate proportional times between the known boundaries
        total_gap_time = next_split_time - prev_split_time
        for i, ratio in enumerate(ratios[:-1]):
            interpolated_split = prev_split_time + round(total_gap_time * sum(ratios[:i + 1]))
            interpolated_splits.append(convert_to_str(interpolated_split))
    elif next_split_time is not None:
        total_gap_time = next_split_time
        for i, ratio in enumerate(ratios[:-1]):
            interpolated_split = round(total_gap_time * sum(ratios[:i + 1]))
            interpolated_splits.append(convert_to_str(interpolated_split))

    return interpolated_splits


def parse_html_into_json(html_filename, bucket_name, parsed_object_filename):

    response = s3.get_object(Bucket=bucket_name, Key=html_filename)
    html_content = response['Body'].read().decode('utf-8')

    title, courses, all_courses_controls, all_courses_runners = parse_raw_html(html_filename, html_content)
    fill_missing_splits_with_proportional_ratio(all_courses_runners)
    splits = calculate_splits_from_cumulative_timestamps(all_courses_runners)

    sorted_split_runners, sorted_cumul_runners = sort_runners_for_each_leg(splits)

    add_split_positions(sorted_split_runners, all_courses_runners)
    add_cumul_positions(sorted_cumul_runners, all_courses_runners)
    add_loss_info(sorted_split_runners, sorted_cumul_runners, splits, all_courses_runners)

    data_json = {'title': title,
                 'courses': courses,
                 'controls': all_courses_controls,
                 'runners': all_courses_runners}

    s3.put_object(Bucket=bucket_name, Key=parsed_object_filename, Body=json.dumps(data_json, indent=2))

    return data_json


def lambda_handler(event, context):
    try:
        print('event:', json.dumps(event))
        # Extract the file name from the query parameters
        html_filename = event['queryStringParameters']['file_to_retrieve']
        # html_filename = 'splits Ineia 27 Oct 2024.html'
        parsed_object_filename = html_filename.replace('.html', '.json')

        # Check if the parsed object file already exists in S3
        bucket_name = 'cyprus-orienteering'

        try:
            s3.head_object(Bucket=bucket_name, Key=parsed_object_filename)
            parsed_object_exists = True
        except s3.exceptions.ClientError as e:
            if e.response['Error']['Code'] == '404':
                parsed_object_exists = False
            else:
                parsed_object_exists = False

        if parsed_object_exists:
            # If parsed object already exists, retrieve it from S3
            response = s3.get_object(Bucket=bucket_name, Key=parsed_object_filename)
            data_json = json.loads(response['Body'].read().decode('utf-8'))
        else:
            data_json = parse_html_into_json(html_filename, bucket_name, parsed_object_filename)

        return {
            'statusCode': 200,
            'body': json.dumps(data_json, indent=2),
            'headers': {
                'Content-Type': 'application/json',
                'Access-Control-Allow-Origin': '*',
            }

        }

    except Exception as e:
        return {
            'statusCode': 500,
            'body': 'Error retrieving file: ' + str(e),
            'headers': {
                'Content-Type': 'application/json',
                'Access-Control-Allow-Origin': '*',
            }
        }
