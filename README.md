# heavy_files

I made this program as answer to a [question on reddit](https://www.reddit.com/r/software/comments/nst42o/is_there_an_app_that_can_make_my_files_harder_to/).

It works with selected folders, links, and multiple files, but unfortunately, I haven't been able to get it to work with files on the desktop.

> usage: heavy_files.py [-h] [--max_size MAX_SIZE] [--min_speed MIN_SPEED] [--max_speed MAX_SPEED] [--refresh_rate REFRESH_RATE]
>
> Slows down the mouse pointer based on the size of the selected files.
>
> optional arguments:
>   -h, --help                   show this help message and exit
>   --max_size MAX_SIZE          Size in MB. The maximum size, when the speed is the minimum. (default: 500 Bytes)
>   --min_speed MIN_SPEED        Speed between 1 and 20. The minimum speed. (default: 1)
>   --max_speed MAX_SPEED        Speed between 1 and 20. The maximum speed. (default: 20 (current speed))
>   --refresh_rate REFRESH_RATE  (default: 1 s)

