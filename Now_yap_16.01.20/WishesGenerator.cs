using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Now_yap_16._01._20
{
    public class WishesGenerator
    {
        private List<List<string>> allWishes;
        private List<List<string>> combinationsWishes;
        private List<List<string>> combinationsWishesNotUnique = null;
        private List<List<string>> combinationsWishesUnique = null;
        private List<string> wishesTrio;
        private int currCount;
        private int maxCountUniqueTrio;
        private Random random;


        public int MaxCountUniqueTrio
        {
            get
            {
                if (maxCountUniqueTrio == -1) countMaxCombination();
                return maxCountUniqueTrio;
            }
        }

        public List<string> WishesTrio
        {
            get
            {
                return new List<string> { (string)wishesTrio[0].Clone(), (string)wishesTrio[1].Clone(), (string)wishesTrio[2].Clone() };
            }
        }

        public WishesGenerator(List<List<string>> allWishes)
        {
            this.allWishes = new List<List<string>>();
            for (int i = 0; i < allWishes.Count; i++)
            {
                this.allWishes.Add(new List<string>());
                for (int j = 0; j < allWishes[i].Count; j++)
                {
                    this.allWishes[i].Add((string)allWishes[i][j].Clone());
                }
            }
            wishesTrio = new List<string>();
            random = new Random();
            maxCountUniqueTrio = -1;
        }

        private void countMaxCombination()
        {
            maxCountUniqueTrio = 0;
            for (int topic0 = 0; topic0 < allWishes.Count; topic0++)
            {
                for (int topic1 = topic0 + 1; topic1 < allWishes.Count; topic1++)
                {
                    for (int topic2 = topic1 + 1; topic2 < allWishes.Count; topic2++)
                    {
                        maxCountUniqueTrio += allWishes[topic0].Count * allWishes[topic1].Count * allWishes[topic2].Count;
                    }
                }
            }
        }

        public bool isThereEnoughtCombination(int countNames)
        {
            return MaxCountUniqueTrio <= countNames;
        }

        public void generateWishes()
        {
            combinationsWishes = new List<List<string>>();
            combinationsWishesUnique = new List<List<string>>();
            combinationsWishesNotUnique = new List<List<string>>();
            bool[] topicsUse = new bool[allWishes.Count];

            for (int topic0 = 0; topic0 < allWishes.Count; topic0++)
            {
                for (int topic1 = topic0 + 1; topic1 < allWishes.Count; topic1++)
                {
                    for (int topic2 = topic1 + 1; topic2 < allWishes.Count; topic2++)
                    {
                        for (int i = 0; i < allWishes[topic0].Count; i++)
                        {
                            for (int j = 0; j < allWishes[topic1].Count; j++)
                            {
                                for (int k = 0; k < allWishes[topic2].Count; k++)
                                {
                                    combinationsWishes.Add(new List<string> { allWishes[topic0][i], allWishes[topic1][j], allWishes[topic2][k] });
                                }
                            }
                        }
                    }
                }
            }

            int tek;
            while (combinationsWishes.Count > 0)
            {
                tek = random.Next() % combinationsWishes.Count;
                if (isUniqueCombination(combinationsWishes[tek]))
                {
                    combinationsWishesUnique.Add(combinationsWishes[tek]);
                }
                else
                {
                    combinationsWishesNotUnique.Add(combinationsWishes[tek]);
                }
                combinationsWishes.RemoveAt(tek);
            }

            combinationsWishes = null;
            currCount = 0;
        }

        public void newTrio()
        {
            if (combinationsWishesUnique != null)
            {
                if (currCount < combinationsWishesUnique.Count)
                {
                    wishesTrio = combinationsWishesUnique[currCount];
                }
                else if (currCount - combinationsWishesUnique.Count < combinationsWishesNotUnique.Count)
                {
                    wishesTrio = combinationsWishesNotUnique[currCount - combinationsWishesUnique.Count];
                }
                else
                {
                    wishesTrio = new List<string> { "#No Wish#", "#No Wish#", "#No Wish#" };
                }
            }
            else
            {
                this.generateWishes();
                this.newTrio();
                return;
            }
            currCount++;
        }

        private bool isUniqueCombination(List<string> comb)
        {
            foreach (var elem in combinationsWishesUnique)
            {
                if (
                    elem[0] == comb[0] || elem[1] == comb[1] || elem[2] == comb[2]
                    || elem[0] == comb[1] || elem[1] == comb[2] || elem[2] == comb[0]
                    || elem[0] == comb[2] || elem[1] == comb[0] || elem[2] == comb[1]
                    )
                    return false;
            }
            return true;
        }

        public void writeAllCombinations()
        {
            if (combinationsWishesUnique != null)
            {
                int oldCur = currCount;
                currCount = 0;
                for (int j = 0; j < MaxCountUniqueTrio; j++)
                {
                    newTrio();
                    for (int i = 0; i < 3; i++)
                    {
                        Console.Write($"{WishesTrio[i]} ");
                    }
                    Console.WriteLine();
                }
                currCount = oldCur;
            }
            else
            {
                generateWishes();
                writeAllCombinations();
                return;
            }
        }
    }
}
