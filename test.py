def removeDuplicates(nums):
        """
        :type nums: List[int]
        :rtype: int
        """
        k = 0
        tmp = { }

        nums_copy = nums.copy()
        
        for number in nums_copy:
            if number not in tmp:
                tmp[number] = True  # tmp{ 1: True }
                k += 1
            else:
                 nums.remove(number)

        print(k)
        return k

nums = [1,1,2]
removeDuplicates(nums)

print(nums)