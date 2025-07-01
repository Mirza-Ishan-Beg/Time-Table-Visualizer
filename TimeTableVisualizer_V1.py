import matplotlib.pyplot as plt  # MAKE SURE TO INSTALL microsoft visual studio distribution 2019 along with "pip install msvc-runtime"
import numpy as np
import pandas as pd
import re

class WeeklySchedulePlotter:
    def __init__(self, excel_file):
        self.df = pd.read_excel(excel_file)
        self.time_picker = r"\d{4}"
        self.block_picker = r'\([^@^/]+\)'
        self.dic_time_lessons = {}
        self.day = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
        self.process_schedule()

    def process_schedule(self):
        for subject in self.df.columns:
            if subject != "DAY IN THE WEEK":
                ptr = 0
                for cell in self.df[subject]:
                    if cell != "NONE":
                        matched_times_list = re.findall(self.time_picker, cell)
                        matched_block = re.findall(self.block_picker, cell + "@")
                        matched_block = " ".join(matched_block)
                        if self.day[ptr] not in self.dic_time_lessons:
                            self.dic_time_lessons[self.day[ptr]] = [{subject + "/" + matched_block: matched_times_list}]
                        else:
                            self.dic_time_lessons[self.day[ptr]].append({subject + "/" + matched_block: matched_times_list})
                    ptr += 1

    def convert_times(self, times_list):
        converted_times = []
        for time_range in times_list:
            if len(time_range) % 2 == 0:
                if len(time_range) > 2:
                    start_time = time_range[0]
                    end_time = time_range[-1]
                else:
                    start_time = time_range[0]
                    end_time = time_range[1]

                start_hour, start_minute = int(start_time[:2]), int(start_time[2:])
                end_hour, end_minute = int(end_time[:2]), int(end_time[2:])
                converted_times.append((start_hour, start_minute, end_hour, end_minute))
            else:
                raise ValueError("Each range should have an even number of elements (pairs of start and end times).")
        return converted_times

    def plot_clock(self, highlighted_times, custom_tags=None, save_path=None, day_index=0):
        fig, ax = plt.subplots(subplot_kw={'projection': 'polar'}, figsize=(10, 10))

        fig.patch.set_facecolor('black')
        ax.set_facecolor('black')

        num_segments = 24 * 60
        theta = np.linspace(0, 2 * np.pi, num_segments, endpoint=False)
        r = np.ones_like(theta)

        ax.plot(theta, r, color='white', linewidth=1)

        for minute in range(0, 60, 5):
            for hour in range(24):
                angle = ((hour * 60 + minute) / num_segments) * 2 * np.pi
                if minute == 0:
                    ax.plot([angle, angle], [0, 1], color='darkgrey', linestyle='-', linewidth=1)
                elif minute % 15 == 0:
                    ax.plot([angle, angle], [0, 1], color='darkgrey', linestyle='--', linewidth=0.4)
                elif minute % 5 == 0:
                    ax.plot([angle, angle], [0, 1], color='darkgrey', linestyle='-', linewidth=0.2)

        prior_start_hour = prior_start_min = prior_end_hour = prior_end_min = None
        offsets = 0.7
        offsetter_midpoint = 0

        for (start_hour, start_minute, end_hour, end_minute), tag_text in zip(highlighted_times, custom_tags or []):
            start_angle = ((start_hour * 60 + start_minute) / num_segments) * 2 * np.pi
            end_angle = ((end_hour * 60 + end_minute) / num_segments) * 2 * np.pi

            if end_angle < start_angle:
                end_angle += 2 * np.pi

            ax.fill_between([start_angle, end_angle], 0, 1, color='skyblue', alpha=0.7)

            midpoint_angle = (start_angle + end_angle) / 2
            if end_angle < start_angle:
                midpoint_angle += np.pi

            if (prior_start_hour, prior_start_min, prior_end_hour, prior_end_min) == (start_hour, start_minute, end_hour, end_minute):
                offsetter_midpoint -= 0.12
                midpoint_angle -= offsetter_midpoint
            else:
                offsetter_midpoint = 0

            label = f'{start_hour:02d}:{start_minute:02d} - {end_hour:02d}:{end_minute:02d}'
            ax.text(midpoint_angle, offsets, label, horizontalalignment='center', verticalalignment='center',
                    fontsize=7, fontweight='bold', alpha=0.8, color='yellow')
            if tag_text:
                ax.text(midpoint_angle-0.06, offsets, tag_text, horizontalalignment='center', verticalalignment='center',
                        fontsize=5, fontweight='bold', alpha=0.9, color='orange')

            prior_start_hour, prior_start_min, prior_end_hour, prior_end_min = start_hour, start_minute, end_hour, end_minute

        for hour in range(24):
            angle = ((hour * 60) / num_segments) * 2 * np.pi
            label = f'{hour:02d}:00'
            ax.text(angle, 1.1, label, horizontalalignment='center', verticalalignment='center',
                    fontsize=12, fontweight='bold', alpha=0.8, color='white')

        ax.set_yticklabels([])
        ax.set_xticklabels([])
        ax.spines['polar'].set_visible(False)
        titlepart = f'24-Hour Clock of {self.day[day_index]}'
        plt.title(titlepart, color='white', pad=20)

        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight', facecolor='black')
        plt.close()

    def generate_plots(self):
        c = 0
        for day, time_blocks in self.dic_time_lessons.items():
            tags = []
            times = []
            for entry in time_blocks:
                tags.append(list(entry.keys())[0])
                times.append(entry[tags[-1]])

            converted_times = self.convert_times(times)
            save_path = f'timetables_chart/high_res_clock_of_{self.day[c]}.png'
            self.plot_clock(converted_times, tags, save_path=save_path, day_index=c)
            c += 1


# Example usage
if __name__ == "__main__":
    xl_file = "Weekly_Schedule.xlsx"
    plotter = WeeklySchedulePlotter(xl_file)
    plotter.generate_plots()
