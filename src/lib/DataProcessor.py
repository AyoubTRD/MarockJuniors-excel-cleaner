class DataProcessor:
    def __init__(self):
        pass

    def rows_are_same_identity(self, row1, row2, unique_fields):
        is_duplicate = True

        for field in unique_fields:
            if row1[field].lower().strip() != row2[field].lower().strip():
                is_duplicate = False
                break

        return is_duplicate

    def row_is_invalid(self, row, required_fields): 
        for field in required_fields:
            if not row[field]:
                return True

        return False

    # data is supposed to be a 2d matrix, with no header row
    # Including the header wouldn't lead to any issues though
    # unique_fields is supposed to represent the indices of rows
    # that should be unique
    def group_duplicates(self, data, unique_fields, required_fields):
        new_data = []

        invalid_rows = []

        for row in data:
            if self.row_is_invalid(row, required_fields):
                invalid_rows.append(row)
                continue

            is_processed_row = False
            for processed_row in new_data:
                # Row has already been processed, move to the next row
                if self.rows_are_same_identity(row, processed_row, unique_fields):
                    is_processed_row = True
                    break
            
            if is_processed_row: continue

            duplicate_rows = []

            for row2 in data:
                if self.rows_are_same_identity(row, row2, unique_fields):
                    duplicate_rows.append(row2)
            
            for row in duplicate_rows:
                new_data.append(row)
        
        for invalid_row in invalid_rows:
            new_data.append(invalid_row)

        return new_data

    def merge_data(self, grouped_data, unique_fields):
        merged_data = []
        current_index = 0

        while current_index < len(grouped_data):
            current_row = grouped_data[current_index]

            rows = [current_row]

            while True:
                next_index = current_index + 1
                if next_index >= len(grouped_data): break
                
                is_duplicate = self.rows_are_same_identity(current_row, grouped_data[next_index], unique_fields)

                if is_duplicate: 
                    rows.append(grouped_data[next_index])
                else: break

                current_index += 1
            
            row = self.merge_rows(rows)
            merged_data.append(row)
            current_index += 1

        return merged_data

    def merge_rows(self, rows):
        if len(rows) == 0: return None

        merged_row = rows[0]

        n_fields = len(rows[0])
        for field in range(n_fields):
            merged_field = rows[0][field]

            for row in rows:
                if row[field]: merged_field = row[field]
            
            merged_row[field] = merged_field

        return merged_row
            

