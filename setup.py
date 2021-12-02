import setuptools

setuptools.setup(
    name="budgetize",
    version="0.1.0",
    author="Ethan D. Twardy",
    author_email="ethan.twardy@gmail.com",
    description="Scripts to help organize and perform budgeting tasks",
    url="https://github.com/AmateurECE/Budget",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
    provides=['budgetize']
)
