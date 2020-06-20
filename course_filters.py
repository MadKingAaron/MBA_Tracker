from course_values import *


# Function Selector
def course_filter_selector(course_name):
    if "500" in course_name:
        return course_500
    elif "502" in course_name:
        return course_502
    elif "504" in course_name:
        return course_504
    elif "506" in course_name:
        return course_506
    elif "508" in course_name:
        return course_508
    elif "510" in course_name:
        return course_510
    elif "520" in course_name:
        return course_520
    elif "521" in course_name:
        return course_521
    elif "523" in course_name:
        return course_523
    elif "524" in course_name:
        return course_524
    elif "525" in course_name:
        return course_525
    elif "526" in course_name:
        return course_526
    elif "530" in course_name:
        return course_530
    elif "531" in course_name:
        return course_531
    elif "540" in course_name:
        return course_540
    elif "541" in course_name:
        return course_541
    elif "545" in course_name:
        return course_545
    elif "550" in course_name:
        return course_550
    elif "555" in course_name:
        return course_555
    elif "560" in course_name:
        return course_560
    elif "570" in course_name:
        return course_570
    elif "572" in course_name:
        return course_572
    else:
        return allow_all


def allow_all(student_classes):
    return True


# Foundation Course Functions

def course_500(student_classes):
    # print("Name: "+student_classes[1])
    # print(student_classes[B500])
    return student_classes[B500] is None


def course_502(student_classes):
    return student_classes[B502] is None


def course_504(student_classes):
    return student_classes[B504] is None


def course_506(student_classes):
    return student_classes[B506] is None


def course_508(student_classes):
    return student_classes[B508] is None


def taken_all_foundation(student_classes):
    if course_500(student_classes) is not True and course_502(student_classes) is not True and course_504(
            student_classes) is not True and course_506(student_classes) is not True and course_508(
            student_classes) is not True:
        return True
    else:
        return False


# Core Course Functions

def count_core_courses(student_classes):
    core_count = 0
    for i in range(B510, B572):
        if student_classes[i] is not None:
            core_count += student_classes[i]

    return core_count


def course_510(student_classes):
    return student_classes[B510] is None and taken_all_foundation(student_classes)


def course_520(student_classes):
    return student_classes[B520] is None and taken_all_foundation(student_classes)


def course_530(student_classes):
    return taken_all_foundation(student_classes) and (student_classes[B530] is None) and (
            student_classes[B510] is not None or student_classes[B520] is not None) and (
                   count_core_courses(student_classes) < 10)


def course_540(student_classes):
    return taken_all_foundation(student_classes) and (student_classes[B540] is None) and (
            student_classes[B510] is not None or student_classes[B520] is not None) and (
                   count_core_courses(student_classes) < 10)


def course_545(student_classes):
    return taken_all_foundation(student_classes) and (student_classes[B545] is None) and (
            student_classes[B510] is not None or student_classes[B520] is not None) and (
                   count_core_courses(student_classes) < 10)


def course_550(student_classes):
    return taken_all_foundation(student_classes) and (student_classes[B550] is None) and (
            student_classes[B510] is not None or student_classes[B520] is not None) and (
                   count_core_courses(student_classes) < 10)


def course_555(student_classes):
    return taken_all_foundation(student_classes) and (student_classes[B555] is None) and (
            student_classes[B510] is not None or student_classes[B520] is not None) and (
                   count_core_courses(student_classes) < 10)


def course_560(student_classes):
    return taken_all_foundation(student_classes) and (student_classes[B560] is None) and (
            student_classes[B510] is not None or student_classes[B520] is not None) and (
                   count_core_courses(student_classes) < 10)


def course_570(student_classes):
    core_count = count_core_courses(student_classes)

    # Temporary
    return (student_classes[B510] is not None) and (student_classes[B520] is not None) and (
                student_classes[B530] is not None) and (student_classes[B540] is not None) and (
                       student_classes[B550] is not None) and (student_classes[B560] is not None) and (
                       student_classes[B570] is None) and (student_classes[3] == "MBA") and (core_count >= 7)

    return ((core_count >= 7) and (core_count < 10)) and student_classes[B570] is None


# Special Topics Functions

def course_521(student_classes):
    return calculate_eligible_for_st(student_classes)


def course_523(student_classes):
    return calculate_eligible_for_st(student_classes)


def course_524(student_classes):
    return calculate_eligible_for_st(student_classes)


def course_525(student_classes):
    return calculate_eligible_for_st(student_classes)


def course_526(student_classes):
    return calculate_eligible_for_st(student_classes)


def course_531(student_classes):
    return calculate_eligible_for_st(student_classes)


def course_541(student_classes):
    return calculate_eligible_for_st(student_classes)


def course_572(student_classes):
    return taken_all_foundation(student_classes) and (count_core_courses(student_classes) >= 2)


def calculate_eligible_for_st(student_classes):
    return (student_classes[B510] is not None and student_classes[B520] is not None) and (
            calculate_total_special_topics(student_classes) < 3)


def calculate_total_special_topics(student_classes):
    total_st = 0
    for i in range(B521, B530):
        if student_classes[i] is not None:
            total_st += student_classes[i]

    # Check for 531
    if student_classes[B531] is not None:
        total_st += student_classes[B531]

    # Check for 541
    if student_classes[B531] is not None:
        total_st += student_classes[B541]

    return total_st
