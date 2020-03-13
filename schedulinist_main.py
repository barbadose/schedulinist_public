import time
import calendar
from datetime import datetime
from random import shuffle
import networkx as nx
import xlsxwriter

# raw_availabilites_mock = [
#     {
#         "name": "אשר",
#         "availabilities": [1, 2, 6,6,6,6,6, 14, 19, 26, 30],
#         "phone": "censored phone number",
#     },
#     {
#         "name": "לירן",
#         "availabilities": [1, 6, 7, 12, 13, 19, 20, 22, 27],
#         "phone": "censored phone number",
#     },
#     {
#         "name": "ינאי",
#         "availabilities": [7, 8, 13, 14, 20, 21, 28],
#         "phone": "censored phone number",
#     },
#     {
#         "name": "עמית",
#         "availabilities": [6, 8, 15, 21, 22, 29],
#         "phone": "censored phone number",
#     },
#     {"name": "יוני", "availabilities": [7, 9, 23, 26], "phone": "censored phone number",},
#     {"name": "מריה", "availabilities": [9, 21, 26, 28, 25], "phone": "censored phone number",},
# ]
# OR_phones_mock = [
#     {"name": "מחלקה", "phone": "censored phone number"},
#     {"name": "חדר צנתורים", "phone": "censored phone number"},
# ]



class Schedulinist:
    """
    >>> def __init__(
        self,
        month_year=default_month_year,
        days_to_remove=[],
        raw_availabilities=None,
        export_path=export_path,
    )
    """

    def __init__(
        self,
        month_year=default_month_year,
        days_to_remove=[],
        raw_availabilities=None,
        export_path=None,
        or_phones = None,
    ):
        self.month = month_year[0]
        self.year = month_year[1]
        self.days_to_remove = days_to_remove
        self.raw_availabilities = raw_availabilities
        self.export_path = export_path
        self.or_phones = or_phones
        if self.raw_availabilities == None:
            raise NameError(
                "A Schedulinist object cannot be instantiated without supplying raw_availabilities in args"
            )
        if self.export_path == None:
            self.export_path = "schedulinist_output"  + "_" + str(self.month) + "." + str(self.year) + "_"  +  str(datetime.now().time().microsecond) + ".xlsx"
        elif self.export_path != None:
            self.export_path = self.export_path + "_" + str(self.month) + "." + str(self.year) + "_"  +  str(datetime.now().time().microsecond) + ".xlsx"
        if self.or_phones == None:
            self.or_phones =  [
                {"name": "מחלקה", "phone": "censored phone number"},
                {"name": "חדר צנתורים", "phone": "censored phone number"},
            ]

    def get_month_work_days(self, month=None, year=None, days_to_remove=None):
        """
        a function that returns legitimate work days of a chosen month.
        parameters:
        ------------ 
        month as int (i.e. 4 for April) \n
        year as int (i.e. 2020) \n
        days_to_remove as list of ints (i.e. [8, 9, 21, 28]) \n
        returns: 
        ---------
        days_to_fill as list of lists of ints (i.e. [[2, 3], [5, 6, 7]])
        """
        if month == None:
            month = self.month
        if year == None:
            year = self.year
        if days_to_remove == None:
            days_to_remove = self.days_to_remove

        days_to_remove.append(0)
        cal = calendar.Calendar()
        cal.setfirstweekday(6)
        selected_month = cal.monthdayscalendar(year, month)

        # Removes fridays and saturdays:
        for week in selected_month:
            week.pop()
            week.pop()
        # print("selected_month", selected_month)

        # Removes 0s and days_to_remove:
        days_to_fill = [
            [day for day in week if day not in days_to_remove]
            for week in selected_month
        ]
        return days_to_fill

    def undesirable_days(self, days_to_fill=None, raw_availabilites=None):
        """
        Returns a list of workdays where no one can work.
        
        Parameters:
        ----------- 
            raw availabilities is a list of dicts in the form:
                >>> raw_availabilites = [
                {"name": "Asher", "availabilities": [1, 6, 14, 19, 26, 30]},
                {"name": "Liran", "availabilities": [1, 6, 7, 13, 19, 20, 22, 27]},
                {"name": "Yanay", "availabilities": [1, 7, 8, 13, 14, 20, 21, 28]},
                {"name": "Amit", "availabilities": [2, 6, 8, 15, 21, 22, 29]},
                {"name": "Yoni", "availabilities": [2, 7, 9, 23, 26]},
                {"name": "Maria", "availabilities": [2, 9, 21, 26, 28]},]
            
            days to fill is a list of ints in the form:
                >>> days_to_fill = [[1, 2], [5, 6, 7], [12, 13, 14, 15, 16], [19, 20, 22, 23], [26, 27, 29, 30]]
        """
        if days_to_fill == None:
            days_to_fill = self.get_month_work_days()
        if raw_availabilites == None:
            raw_availabilites = self.raw_availabilities

        unique_availabilities = list(
            set(
                [
                    day
                    for linist in raw_availabilites
                    for day in linist["availabilities"]
                ]
            )
        )
        undesirable_days = [
            day
            for week in days_to_fill
            for day in week
            if day not in unique_availabilities
        ]

        return undesirable_days

    def desirable_days(self, days_to_fill=None, raw_availabilites=None):
        """
        Returns a list of workdays at least one linist can work on.
        
        Parameters:
        ----------- 
            raw availabilities is a list of dicts in the form:
                >>> raw_availabilites = [
                {"name": "Asher", "availabilities": [1, 6, 14, 19, 26, 30]},
                {"name": "Liran", "availabilities": [1, 6, 7, 13, 19, 20, 22, 27]},
                {"name": "Yanay", "availabilities": [1, 7, 8, 13, 14, 20, 21, 28]},
                {"name": "Amit", "availabilities": [2, 6, 8, 15, 21, 22, 29]},
                {"name": "Yoni", "availabilities": [2, 7, 9, 23, 26]},
                {"name": "Maria", "availabilities": [2, 9, 21, 26, 28]},]
            
            days to fill is a list of ints in the form:
                >>> days_to_fill = [[1, 2], [5, 6, 7], [12, 13, 14, 15, 16], [19, 20, 22, 23], [26, 27, 29, 30]]
        """

        if days_to_fill == None:
            days_to_fill = self.get_month_work_days()
        if raw_availabilites == None:
            raw_availabilites = self.raw_availabilities

        unique_availabilities = list(
            set(
                [
                    day
                    for linist in raw_availabilites
                    for day in linist["availabilities"]
                ]
            )
        )
        desirable_days = [
            day for week in days_to_fill for day in week if day in unique_availabilities
        ]
        return desirable_days

    def get_clean_month(self, days_to_fill=None, raw_availabilites=None):
        """
        a function that returns days_to_fill without undesirable days.
        parameters:
        ------------ 
        days_to_fill as a list of lists of ints (i.e. [[1, 2], [5, 6, 7]]) \n
        raw_availabilities \n
        returns: 
        ----------
        clean_month as a list of lists of ints (i.e. [[1], [5, 7]])
        """
        if days_to_fill == None:
            days_to_fill = self.get_month_work_days()
        if raw_availabilites == None:
            raw_availabilites = self.raw_availabilities

        clean_month = [
            [
                day
                for day in week
                if day not in (self.undesirable_days(days_to_fill, raw_availabilites))
            ]
            for week in days_to_fill
        ]
        return clean_month

    def single_week_maxflow(
        self, clean_week, raw_availabilites, weights, s_edge_capacity=1
    ):
        """
        "clean_week" must not include undesirable days.
        """
        DG = nx.DiGraph()
        week_length = len(clean_week)
        edge_capacity = 1
        s_weight = 1
        t_weight = 1
        growing_edges_list = []
        # Add all linist edges and source edges:
        for linist in raw_availabilites:
            linist_name = linist["name"]
            linist_weight = weights[linist_name]
            linist_availabilities = [
                day for day in linist["availabilities"] if day in clean_week
            ]
            new_s_edge_tuple = (
                "s",
                linist_name,
                {"capacity": s_edge_capacity, "weight": s_weight},
            )
            growing_edges_list.append(new_s_edge_tuple)
            for availability in linist_availabilities:
                new_edge_tuple = (
                    linist_name,
                    availability,
                    {"capacity": edge_capacity, "weight": linist_weight},
                )
                growing_edges_list.append(new_edge_tuple)
        # Add all sink (t) edges:
        for day in clean_week:
            new_t_edge_tuple = (
                day,
                "t",
                {"capacity": edge_capacity, "weight": t_weight},
            )
            growing_edges_list.append(new_t_edge_tuple)
        # Run max flow min cost:
        
        DG.add_edges_from(growing_edges_list)
        max_flow_response = nx.algorithms.max_flow_min_cost(DG, ("s"), ("t"))
        filled_days_count = sum(max_flow_response["s"].values())
        if filled_days_count == week_length:
            return max_flow_response
        elif filled_days_count < week_length:
            return self.single_week_maxflow(
                clean_week, raw_availabilites, weights, s_edge_capacity + 1
            )
        else:
            raise NameError(
                "filled_days_count exceeded week_length, that's really weird."
            )

    def month_maxflow(self, clean_month=None, raw_availabilites=None):

        """
        takes clean_month and raw_availabilities
        returns final_result i.e.:
        >>> {'s': {'Yoni': 3, 'Asher': 3, 'Amit': 3, 'Liran': 5, 'Maria': 0, 'Yanay': 2}, 'Yoni': {23: 1, 7: 1, 26: 1}, 'Asher': {14: 1, 19: 0, 6: 0, 2: 1, 1: 0, 30: 1, 26: 0}, 'Amit': {15: 1, 22: 1, 6: 0, 29: 1}, 'Liran': {13: 0, 12: 1, 19: 1, 20: 0, 22: 0, 7: 0, 6: 1, 1: 1, 27: 1}, 'Maria': {26: 0}, 'Yanay': {13: 1, 14: 0, 20: 1, 7: 0}}

        clean_month must not include undesirable days.
        """
        if clean_month == None:
            clean_month = self.get_clean_month()
        if raw_availabilites == None:
            raw_availabilites = self.raw_availabilities

        # Shuffle clean_month and raw_availabilities
        for week in clean_month:
            shuffle(week)
        shuffle(clean_month)
        for linist in raw_availabilites:
            shuffle(linist["availabilities"])
        shuffle(raw_availabilites)
        # Initialize weights_dict:
        weights_dict = {}
        for linist in raw_availabilites:
            weights_dict[linist["name"]] = 1
        response_list = []
        # Run all weeks in order and mutate response_list accordingly:
        for clean_week in clean_month:
            # if clean_week == [] (no one can work on this week):
            if not clean_week:
                continue
            else:
                response = self.single_week_maxflow(
                    clean_week, raw_availabilites, weights_dict
                )
                response_list.append(response)
                for linist, weight in weights_dict.items():
                    weight = weight * (2 ^ (response["s"][linist]))
        # Parse response_list and return output_dict:
        # {'s': {'Asher': 4, 'Liran': 4, 'Yanay': 4, 'Amit': 3, 'Yoni': 1, 'Maria': 1},
        # 'Asher': {1: 1, 2: 1}, 'Liran': {3: 1, 4: 1}, 'Yanay': {5: 1}, 'Amit': {6: 1},
        # 'Yoni': {7: 1}, 'Maria': {8: 1}}
        final_result = {"s": {}}
        for linist in raw_availabilites:
            final_result[linist["name"]] = {}
            final_result["s"][linist["name"]] = 0
        for response in response_list:
            for linist_name, assign_count in response["s"].items():
                final_result["s"][linist_name] = (
                    final_result["s"][linist_name] + assign_count
                )
                final_result[linist_name].update(response[linist_name])

        return final_result

    def export_excel(
        self,
        month=None,
        year=None,
        export_path=None,
        days_to_remove=None,
        raw_availabilites=None,
        or_phones=None,
        undesirable_days=None,
        output=None,
    ):
        """
        exports a Schedulinist instance to excel file.
        """
        if month == None:
            month = self.month
        if year == None:
            year = self.year
        if export_path == None:
            export_path = self.export_path
        if days_to_remove == None:
            days_to_remove = self.days_to_remove
        if raw_availabilites == None:
            raw_availabilites = self.raw_availabilities
        if or_phones == None:
            or_phones = self.or_phones
        if undesirable_days == None:
            undesirable_days = self.undesirable_days()
        if output == None:
            output = self.month_maxflow()

        # get whole month with edges from previous + next months.
        cal = calendar.Calendar()
        cal.setfirstweekday(6)
        selected_month = cal.monthdayscalendar(year, month)

        # Removes fridays and saturdays:
        for week in selected_month:
            week.pop()
            week.pop()
        # Initialzie unassigned desirables list:
        unassigned_desirables = {}
        for linist in output["s"]:
            unassigned_desirables[linist] = []
        # Creates a workbook and adds a worksheet (i.e. April 2020)
        worksheet_name = calendar.month_name[month] + " " + str(year)
        workbook = xlsxwriter.Workbook(export_path)
        worksheet = workbook.add_worksheet(worksheet_name)
        row = 1
        col = 1
        # ~ Formatting and styling ~
        worksheet.right_to_left()
        col_width = 20.0
        row_height_small = 15
        row_height_medium = 28.5
        row_height_big = 40
            # ~ Colors ~
        default_color = "#000000"
        gray_removed_day = "#808080"
        dark_blue = "#8DB4E2"
        light_blue = "#B8CCE4"
            # ~ Header ~
        header_right_to_left_with_border = workbook.add_format({"reading_order": 2, "font_size": 22, "align": "center", "valign": "vcenter",  "border": True, "top": False, "fg_color": "white"})
        header_right_to_left_dark_blue_with_border = workbook.add_format({"reading_order": 2, "font_size": 22, "align": "center", "valign": "vcenter", "border": True, "bg_color": dark_blue,})
            # ~ Misc. Headers ~
        header_diff_month = workbook.add_format({"reading_order": 2, "font_size": 22, "align": "center", "valign": "vcenter",  "border": True, "fg_color": "white", "diag_border": 1, "diag_type": 3})
        header_unassigned_day = workbook.add_format({"reading_order": 2, "font_size": 11, "align": "center", "valign": "vcenter",  "border": True, "top": False, "bg_color": "red"})
        header_removed_day = workbook.add_format({"reading_order": 2, "font_size": 22, "align": "center", "valign": "vcenter",  "border": True, "top": False, "bg_color": gray_removed_day })
            # ~ Text ~
        text_right_to_left = workbook.add_format({"reading_order": 2, "font_size": 11, "align": "center", "valign": "vcenter", })
        text_right_to_left_light_blue_with_border = workbook.add_format({"reading_order": 2, "font_size": 11, "align": "center", "valign": "vcenter", "bg_color": light_blue, "border": True, "bottom": False})
        text_right_to_left_light_blue_with_rl_border = workbook.add_format({"reading_order": 2, "font_size": 11, "align": "center", "valign": "vcenter", "bg_color": light_blue, "right": True, "left": True})
        text_right_to_left_with_brl_border = workbook.add_format({"reading_order": 2, "font_size": 11, "align": "center", "valign": "vcenter", "border": True, "top": False})
                # ~ Bold Text ~
        bold_text_right_to_left = workbook.add_format({"reading_order": 2, "font_size": 11, "bold": True, "align": "center", "valign": "vcenter", "fg_color": "white"})
        bold_text_right_to_left_dark_blue_with_border = workbook.add_format({"reading_order": 2, "font_size": 11, "bold": True, "align": "center", "valign": "vcenter", "bg_color": dark_blue, "border": True, "bottom": False})
            # ~ Date ~
        date_left_to_right_with_border = workbook.add_format({"reading_order": 1, "font_size": 11, "align": "right", "valign": "vcenter", "border": True, "bottom": False, "fg_color": "white" })
        date_unassigned_day = workbook.add_format({"reading_order": 1, "font_size": 11, "align": "right", "valign": "vcenter", "border": True, "bottom": False, "bg_color": "red" })
        date_removed_day = workbook.add_format({"reading_order": 1, "font_size": 11, "align": "right", "valign": "vcenter", "border": True, "bottom": False, "bg_color": gray_removed_day })

                # create sunday through thursday header and fill out all dates
        HEADER = ["ראשון", "שני", "שלישי", "רביעי", "חמישי"]
        for day_name in HEADER:
            worksheet.write(row, col, day_name, header_right_to_left_dark_blue_with_border)
            col += 1
        worksheet.set_row(row, row_height_medium)
        row = 2
        col = 1
        # Loop through selected_month and write to table:
        for week in selected_month:
            worksheet.set_row(row, row_height_small)
            worksheet.set_row(row + 1, row_height_big)
            for day in week:
                date_str = str(day) + "." + str(month) + "." + str(year)
                worksheet.write(row, col, date_str, date_left_to_right_with_border
    )
                # Different month days:
                if day == 0:
                    # rewrite date:
                    worksheet.write(row, col, "", date_left_to_right_with_border)
                    worksheet.write(row + 1, col, "", header_diff_month)
                    col += 1
                # Manually removed days:
                elif day in days_to_remove:
                    worksheet.write(row, col, date_str, date_removed_day)
                    worksheet.write(row + 1, col, "", header_removed_day)
                    col += 1
                # Undesirable days:
                elif day in undesirable_days:
                    worksheet.write(row, col, date_str, date_unassigned_day)
                    worksheet.write(
                        row + 1,
                        col,
                        "יום ללא ביקוש", header_unassigned_day
                    )
                    col += 1
                # Assigned days:
                else:
                    for linist in output["s"]:
                        if day in output[linist]:
                            if output[linist][day] == 0:
                                unassigned_desirables[linist].append(day)
                            else:  # == 1:
                                worksheet.write(
                                    row + 1, col, linist, header_right_to_left_with_border
                                )
                                col += 1
                        else:
                            continue
            col = 1
            row += 2
        # Write fixed phone numbers:
        col = 1

        for linist in raw_availabilites:
            worksheet.set_column(col, col, col_width)
            worksheet.set_row(row, row_height_small)
            worksheet.set_row(row + 1, row_height_small)
            worksheet.set_row(row + 2, row_height_small)
            worksheet.write(row, col, linist["name"], bold_text_right_to_left_dark_blue_with_border)
            worksheet.write(row + 1, col, linist["phone"], text_right_to_left_light_blue_with_border)
            worksheet.write(row + 2, col, "", text_right_to_left_light_blue_with_rl_border)
            worksheet.write(row + 3, col, "", text_right_to_left_with_brl_border)
            col += 1
        col = 1
        row += 5
        for room in or_phones:
            worksheet.write(row, col, room["name"], bold_text_right_to_left_dark_blue_with_border)
            worksheet.write(row + 1, col, room["phone"], text_right_to_left_with_brl_border)
            col += 1
        col = 1
        row += 8
        worksheet.write(row, col, 'ימים ספייר:')
        row += 2
        for linist, lst in unassigned_desirables.items():
            col = 1
            worksheet.write(row, col, linist, text_right_to_left)
            col += 1
            for day in lst:
                worksheet.write(row, col, day, text_right_to_left)
                col += 1
            row += 1

        workbook.close()




